<?php

namespace App\Services;

use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use RuntimeException;

/**
 * Mengubah 3 file Excel (StandardReport mesin fingerprint, Keterangan_Lain,
 * template_gaji) menjadi 1 file rekap absensi — TANPA database, murni
 * diproses di memori per request lalu langsung di-stream sebagai download.
 *
 * RUANG LINGKUP (sesuai keputusan user):
 * - Hanya rekap absensi. Kolom identitas penggajian di template_gaji.xlsx
 *   (No Rekening, Jabatan, Area/Outlet, Gaji Pokok, Bonus) SENGAJA tidak
 *   diisi otomatis karena tidak ada sumber datanya — dikosongkan, silakan
 *   dilengkapi manual oleh user setelah file diunduh.
 * - Formula bawaan template (kolom L,N,P,R,T,V,X,H) TIDAK disentuh, hanya
 *   kolom tanggal (K,M,O,Q,S,U,W) dan Catatan (Y) yang diisi.
 *
 * ASUMSI YANG PERLU DIKONFIRMASI / DISESUAIKAN (lihat CODE_MAP di bawah):
 * kode di Keterangan_Lain.xlsx (s, sn, c, i, l, a, ...) — arti persisnya
 * ditebak dari contoh data. Ubah CODE_MAP jika arti berbeda di perusahaan
 * Anda, tidak perlu mengubah logika lain di kelas ini.
 */
class AbsensiRekapService
{
    /** Kode pada sheet "SCIL" (Keterangan_Lain.xlsx) -> kategori internal. */
    private const CODE_MAP = [
        's' => 'sakit_skd',      // Sakit dengan Surat Keterangan Dokter
        'sn' => 'sakit_non_skd',  // Sakit tanpa Surat Keterangan Dokter
        'c' => 'cuti',
        'i' => 'izin',
        'l' => 'libur',
        'a' => 'alpa',           // alpa yang sudah ditandai manual oleh admin
        'ih' => 'izin_setengah',  // izin setengah hari (kode belum terverifikasi, sesuaikan bila beda)
    ];

    /** kategori -> kolom "Tanggal ..." di template_gaji.xlsx (kolom "Jumlah ..." sebelahnya sudah berisi formula, tidak disentuh). */
    private const TEMPLATE_DATE_COLUMN = [
        'alpa' => 'K',
        'libur' => 'M',
        'izin' => 'O',
        'sakit_skd' => 'Q',
        'sakit_non_skd' => 'S',
        'cuti' => 'U',
        'izin_setengah' => 'W',
    ];

    private const FIRST_DATA_ROW = 3; // baris pertama data karyawan di template_gaji.xlsx

    public function generate(string $standardReportPath, string $keteranganLainPath, string $templatePath): Spreadsheet
    {
        $karyawan = $this->parseStandardReport($standardReportPath);
        $keterangan = $this->parseKeteranganLain($keteranganLainPath);
        $rekap = $this->gabungkan($karyawan, $keterangan);

        return $this->isiTemplate($templatePath, $rekap);
    }

    /**
     * @return array<string, array{id:string, nama:string, kosong_absen: string[]}>
     *                                                                              keyed by ID karyawan. "kosong_absen" = tanggal yang menurut jadwal
     *                                                                              adalah hari kerja normal, tapi tidak ada satu pun log jam masuk/pulang.
     */
    private function parseStandardReport(string $path): array
    {
        $spreadsheet = IOFactory::load($path);

        $jadwalSheet = $spreadsheet->getSheetByName('Jadwal Info');
        $logSheet = $spreadsheet->getSheetByName('Lap. Log Absen');

        if (! $jadwalSheet || ! $logSheet) {
            throw new RuntimeException(
                'Format 1_StandardReport tidak dikenali: sheet "Jadwal Info" dan/atau "Lap. Log Absen" tidak ditemukan.'
            );
        }

        // --- Sheet "Jadwal Info": ID (A), Nama (B), Departemen (C), lalu tanggal mulai kolom D.
        $tanggalJadwal = $this->bacaHeaderTanggal($jadwalSheet, 3, 4);

        $karyawan = [];
        $row = 5;
        while (true) {
            $id = trim((string) $jadwalSheet->getCell('A'.$row)->getValue());
            if ($id === '') {
                break;
            }
            $nama = trim((string) $jadwalSheet->getCell('B'.$row)->getValue());

            $jadwalPerTanggal = [];
            foreach ($tanggalJadwal as $tanggal => $col) {
                $jadwalPerTanggal[$tanggal] = $jadwalSheet->getCell($col.$row)->getValue();
            }

            $karyawan[$id] = [
                'id' => $id,
                'nama' => $nama,
                'jadwal' => $jadwalPerTanggal, // 1 = hari kerja normal; kosong/lain = bukan hari kerja normal
                'kosong_absen' => [],
                'tidak_finger_pagi' => [],
                'tidak_finger_sore' => [],
            ];
            $row++;
        }

        // --- Sheet "Lap. Log Absen": blok per-karyawan (baris "ID:" lalu baris jam di bawahnya).
        $tanggalLog = $this->bacaHeaderTanggal($logSheet, 4, 1);
        $highestRow = $logSheet->getHighestRow();

        for ($r = 1; $r <= $highestRow; $r++) {
            if (trim((string) $logSheet->getCell('A'.$r)->getValue()) !== 'ID:') {
                continue;
            }
            $id = trim((string) $logSheet->getCell('C'.$r)->getValue());
            $dataRow = $r + 1;

            if (! isset($karyawan[$id])) {
                continue; // ID di log tidak ada di Jadwal Info, lewati
            }

            foreach ($tanggalLog as $tanggal => $col) {
                $jam = trim((string) $logSheet->getCell($col.$dataRow)->getValue());
                $hariKerjaNormal = in_array($karyawan[$id]['jadwal'][$tanggal] ?? null, [1, 1.0, '1'], true);

                if (! $hariKerjaNormal) {
                    continue;
                }

                if ($jam === '') {
                    $karyawan[$id]['kosong_absen'][] = $tanggal;

                    continue;
                }

                // Kolom jam digabung tanpa pemisah antar sesi, mis. "07:5316:02" =
                // sesi 1 "07:53" (pagi) + sesi 2 "16:02" (sore). Ambil per 5 karakter,
                // jam < 12 dianggap sesi pagi, >= 12 dianggap sesi sore.
                [$adaPagi, $adaSore] = $this->klasifikasiJam($jam);

                if (! $adaPagi) {
                    $karyawan[$id]['tidak_finger_pagi'][] = $tanggal;
                }
                if (! $adaSore) {
                    $karyawan[$id]['tidak_finger_sore'][] = $tanggal;
                }
            }
        }

        return $karyawan;
    }

    /**
     * Pecah string jam gabungan (mis. "07:5316:02") jadi potongan 5 karakter
     * ("HH:MM"), lalu tandai apakah ada minimal 1 sesi pagi (jam < 12) dan/atau
     * minimal 1 sesi sore (jam >= 12).
     *
     * @return array{0: bool, 1: bool} [ada_pagi, ada_sore]
     */
    private function klasifikasiJam(string $jamRaw): array
    {
        $adaPagi = false;
        $adaSore = false;
        $panjang = strlen($jamRaw);

        for ($i = 0; $i + 5 <= $panjang; $i += 5) {
            $potongan = substr($jamRaw, $i, 5); // format "HH:MM"
            $jamAngka = (int) substr($potongan, 0, 2);

            if ($jamAngka < 12) {
                $adaPagi = true;
            } else {
                $adaSore = true;
            }
        }

        return [$adaPagi, $adaSore];
    }

    /**
     * @return array<string, array<string, string[]>>
     *                                                [id_karyawan][kategori] = daftar tanggal (angka, sesuai kolom header sheet SCIL)
     */
    private function parseKeteranganLain(string $path): array
    {
        $spreadsheet = IOFactory::load($path);
        $sheet = $spreadsheet->getSheetByName('SCIL') ?? $spreadsheet->getSheet(0);

        $tanggalCol = $this->bacaHeaderTanggal($sheet, 4, 3);

        $hasil = [];
        $row = 5;
        while (true) {
            $id = trim((string) $sheet->getCell('A'.$row)->getValue());
            if ($id === '') {
                break;
            }

            foreach ($tanggalCol as $tanggal => $col) {
                $kodeMentah = strtolower(trim((string) $sheet->getCell($col.$row)->getValue()));
                if ($kodeMentah === '') {
                    continue;
                }
                $kategori = self::CODE_MAP[$kodeMentah] ?? null;
                if ($kategori === null) {
                    // kode tidak dikenal -> tetap dicatat sebagai "lainnya" supaya tidak hilang diam-diam
                    $kategori = 'lainnya_'.$kodeMentah;
                }
                $hasil[$id][$kategori][] = $tanggal;
            }
            $row++;
        }

        return $hasil;
    }

    /**
     * Gabungkan hasil StandardReport + Keterangan_Lain jadi satu rekap final per karyawan.
     */
    private function gabungkan(array $karyawan, array $keterangan): array
    {
        $rekap = [];

        foreach ($karyawan as $id => $data) {
            $kategoriDariKeterangan = $keterangan[$id] ?? [];

            // tanggal yang SUDAH dijelaskan oleh Keterangan_Lain (selain kode alpa 'a')
            $sudahDijelaskan = [];
            foreach ($kategoriDariKeterangan as $kategori => $tanggalList) {
                if ($kategori === 'alpa') {
                    continue;
                }
                $sudahDijelaskan = array_merge($sudahDijelaskan, $tanggalList);
            }

            $alpaOtomatis = array_diff($data['kosong_absen'], $sudahDijelaskan);
            $alpaManual = $kategoriDariKeterangan['alpa'] ?? [];
            $alpa = array_values(array_unique(array_merge($alpaOtomatis, $alpaManual)));
            sort($alpa, SORT_NUMERIC);

            $rekap[$id] = [
                'id' => $id,
                'nama' => $data['nama'],
                'kategori' => array_merge($kategoriDariKeterangan, ['alpa' => $alpa]),
                'tidak_finger_pagi' => $data['tidak_finger_pagi'] ?? [],
                'tidak_finger_sore' => $data['tidak_finger_sore'] ?? [],
            ];
        }

        return $rekap;
    }

    private function isiTemplate(string $templatePath, array $rekap): Spreadsheet
    {
        $spreadsheet = IOFactory::load($templatePath);
        $sheet = $spreadsheet->getSheet(0);

        $row = self::FIRST_DATA_ROW;
        $no = 1;

        foreach ($rekap as $data) {
            $sheet->setCellValue('A'.$row, $no);
            $sheet->setCellValue('B'.$row, $data['nama']);
            // C, D, E, F, G (No Rekening, Jabatan, Area, Gaji Pokok, Bonus) SENGAJA dikosongkan.

            foreach (self::TEMPLATE_DATE_COLUMN as $kategori => $kolom) {
                $tanggalList = $data['kategori'][$kategori] ?? [];
                sort($tanggalList, SORT_NUMERIC);
                $sheet->setCellValue($kolom.$row, implode(',', $tanggalList));
            }

            $sheet->setCellValue('Y'.$row, $this->buatCatatan(
                $data['kategori'],
                $data['tidak_finger_pagi'],
                $data['tidak_finger_sore']
            ));

            $no++;
            $row++;
        }

        // Kosongkan baris template yang tersisa (tidak terpakai) supaya tidak
        // muncul baris "0 Hari" tanpa nama.
        $lastTemplateRow = $sheet->getHighestRow();
        for ($r = $row; $r <= $lastTemplateRow; $r++) {
            foreach (['A', 'B', 'K', 'M', 'O', 'Q', 'S', 'U', 'W', 'Y'] as $kolom) {
                $sheet->setCellValue($kolom.$r, null);
            }
        }

        return $spreadsheet;
    }

    /**
     * Draft catatan otomatis: kategori dari Keterangan_Lain + deteksi
     * "tidak absen pagi/sore" dari pola jam di Lap. Log Absen. Ini tetap
     * DRAFT — hal-hal seperti "terlambat sekian menit", "SP3", "resign"
     * tidak tercakup — cek ulang sebelum dikirim ke payroll.
     *
     * @param  string[]  $tidakFingerPagi  tanggal yang tidak ada sesi absen pagi (jam < 12)
     * @param  string[]  $tidakFingerSore  tanggal yang tidak ada sesi absen sore (jam >= 12)
     */
    private function buatCatatan(array $kategori, array $tidakFingerPagi = [], array $tidakFingerSore = []): string
    {
        $label = [
            'sakit_skd' => 'sakit (SKD)',
            'sakit_non_skd' => 'sakit (non-SKD)',
            'cuti' => 'cuti',
            'izin' => 'izin',
            'izin_setengah' => 'izin setengah hari',
            'libur' => 'libur',
        ];

        $bagian = [];
        foreach ($label as $kategoriKey => $teks) {
            $tanggalList = $kategori[$kategoriKey] ?? [];
            if (empty($tanggalList)) {
                continue;
            }
            sort($tanggalList, SORT_NUMERIC);
            $bagian[] = 'tgl '.implode(', ', $tanggalList).' '.$teks;
        }

        foreach ($kategori as $kategoriKey => $tanggalList) {
            if (! str_starts_with($kategoriKey, 'lainnya_') || empty($tanggalList)) {
                continue;
            }
            sort($tanggalList, SORT_NUMERIC);
            $kode = substr($kategoriKey, strlen('lainnya_'));
            $bagian[] = 'tgl '.implode(', ', $tanggalList)." kode '{$kode}' (kode tidak dikenal, cek CODE_MAP)";
        }

        if (! empty($tidakFingerPagi)) {
            sort($tidakFingerPagi, SORT_NUMERIC);
            $bagian[] = 'tgl '.implode(', ', $tidakFingerPagi).' tidak absen pagi';
        }

        if (! empty($tidakFingerSore)) {
            sort($tidakFingerSore, SORT_NUMERIC);
            $bagian[] = 'tgl '.implode(', ', $tidakFingerSore).' tidak absen sore';
        }

        return implode('; ', $bagian);
    }

    /**
     * Cari kolom-kolom yang berisi header tanggal (angka 1-31) pada suatu
     * baris, mulai dari kolom tertentu, sampai kolom kosong berturut-turut.
     *
     * @return array<int|string, string> [tanggal => kolom_excel]
     */
    private function bacaHeaderTanggal(Worksheet $sheet, int $headerRow, int $startColumnIndex): array
    {
        $hasil = [];
        $kolom = $startColumnIndex;
        $kosongBerturut = 0;

        while ($kosongBerturut < 3) {
            $colLetter = Coordinate::stringFromColumnIndex($kolom);
            $value = $sheet->getCell($colLetter.$headerRow)->getValue();

            if (is_numeric($value) && (int) $value >= 1 && (int) $value <= 31) {
                $hasil[(string) (int) $value] = $colLetter;
                $kosongBerturut = 0;
            } else {
                $kosongBerturut++;
            }
            $kolom++;

            if ($kolom > $startColumnIndex + 60) {
                break; // pengaman, seharusnya tidak pernah tercapai
            }
        }

        return $hasil;
    }
}
