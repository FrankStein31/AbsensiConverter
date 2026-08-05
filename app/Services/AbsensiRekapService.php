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
 * ATURAN LIBUR:
 * - Libur ditentukan dari jadwal (Jadwal Info) pada StandardReport, BUKAN
 *   dari kode manual di Keterangan_Lain. Jika suatu tanggal menurut jadwal
 *   bukan hari kerja normal, tanggal itu otomatis dikategorikan "libur" dan
 *   MENIMPA kode apapun (s/sn/c/i/is) yang mungkin salah diisi di
 *   Keterangan_Lain untuk tanggal tsb.
 *
 * FORMAT TANGGAL:
 * - Semua tanggal disimpan sebagai "d/m" (contoh: "26/5", "1/6") agar jelas
 *   bulannya dan bisa diurutkan kronologis lintas bulan.
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
        'is' => 'izin_setengah',  // izin setengah hari
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

        // --- Ambil bulan awal dari info periode di sheet (biasanya baris 3, misal "2026-05-26 ~ 2026-06-25")
        $bulanAwalJadwal = $this->deteksiBulanAwal($jadwalSheet, 3);
        $bulanAwalLog = $this->deteksiBulanAwal($logSheet, 3);

        // --- Sheet "Jadwal Info": ID (A), Nama (B), Departemen (C), lalu tanggal mulai kolom D.
        $tanggalJadwal = $this->bacaHeaderTanggal($jadwalSheet, 3, 4, $bulanAwalJadwal);

        $karyawan = [];
        $row = 5;
        while (true) {
            $id = trim((string) $jadwalSheet->getCell('A'.$row)->getValue());
            if ($id === '') {
                break;
            }
            $nama = trim((string) $jadwalSheet->getCell('B'.$row)->getValue());

            $jadwalPerTanggal = [];
            foreach ($tanggalJadwal as $tanggalKey => $col) {
                $jadwalPerTanggal[$tanggalKey] = $jadwalSheet->getCell($col.$row)->getValue();
            }

            $karyawan[$id] = [
                'id' => $id,
                'nama' => $nama,
                'jadwal' => $jadwalPerTanggal, // 1 = hari kerja normal; kosong/lain = bukan hari kerja normal (libur)
                'kosong_absen' => [],
                'tidak_finger_pagi' => [],
                'tidak_finger_sore' => [],
            ];
            $row++;
        }

        // --- Sheet "Lap. Log Absen": blok per-karyawan (baris "ID:" lalu baris jam di bawahnya).
        $tanggalLog = $this->bacaHeaderTanggal($logSheet, 4, 1, $bulanAwalLog);
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

            foreach ($tanggalLog as $tanggalKey => $col) {
                $jam = trim((string) $logSheet->getCell($col.$dataRow)->getValue());
                $hariKerjaNormal = in_array($karyawan[$id]['jadwal'][$tanggalKey] ?? null, [1, 1.0, '1'], true);

                if (! $hariKerjaNormal) {
                    continue;
                }

                if ($jam === '') {
                    $karyawan[$id]['kosong_absen'][] = $tanggalKey;

                    continue;
                }

                // Kolom jam digabung tanpa pemisah antar sesi, mis. "07:5316:02" =
                // sesi 1 "07:53" (pagi) + sesi 2 "16:02" (sore). Ambil per 5 karakter,
                // jam < 12 dianggap sesi pagi, >= 12 dianggap sesi sore. Bisa ada lebih
                // dari 1 sesi pagi/sore (double finger) — untuk deteksi ada/tidaknya
                // sesi, cukup ditandai true begitu ketemu minimal satu.
                [$adaPagi, $adaSore] = $this->klasifikasiJam($jam);

                if (! $adaPagi) {
                    $karyawan[$id]['tidak_finger_pagi'][] = $tanggalKey;
                }
                if (! $adaSore) {
                    $karyawan[$id]['tidak_finger_sore'][] = $tanggalKey;
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
     *                                                [id_karyawan][kategori] = daftar tanggal (format "d/m")
     */
    private function parseKeteranganLain(string $path): array
    {
        $spreadsheet = IOFactory::load($path);
        $sheet = $spreadsheet->getSheetByName('SCIL') ?? $spreadsheet->getSheet(0);

        $bulanAwal = $this->deteksiBulanAwal($sheet, 3);
        $tanggalCol = $this->bacaHeaderTanggal($sheet, 4, 3, $bulanAwal);

        $hasil = [];
        $row = 5;
        while (true) {
            $id = trim((string) $sheet->getCell('A'.$row)->getValue());
            if ($id === '') {
                break;
            }

            foreach ($tanggalCol as $tanggalKey => $col) {
                $kodeMentah = strtolower(trim((string) $sheet->getCell($col.$row)->getValue()));
                if ($kodeMentah === '') {
                    continue;
                }
                $kategori = self::CODE_MAP[$kodeMentah] ?? null;
                if ($kategori === null) {
                    // kode tidak dikenal -> tetap dicatat sebagai "lainnya" supaya tidak hilang diam-diam
                    $kategori = 'lainnya_'.$kodeMentah;
                }
                $hasil[$id][$kategori][] = $tanggalKey;
            }
            $row++;
        }

        return $hasil;
    }

    /**
     * Gabungkan hasil StandardReport + Keterangan_Lain jadi satu rekap final per karyawan.
     *
     * Libur dihitung dari jadwal (bukan dari kode manual), dan menimpa kategori
     * lain di tanggal yang sama apabila ada kode yang salah diisi pada hari libur.
     */
    private function gabungkan(array $karyawan, array $keterangan): array
    {
        $rekap = [];

        foreach ($karyawan as $id => $data) {
            $kategoriDariKeterangan = $keterangan[$id] ?? [];
            $jadwal = $data['jadwal'] ?? [];

            // Tanggal yang menurut jadwal BUKAN hari kerja normal -> otomatis libur.
            $liburDariJadwal = [];
            foreach ($jadwal as $tanggal => $nilai) {
                $hariKerjaNormal = in_array($nilai, [1, 1.0, '1'], true);
                if (! $hariKerjaNormal) {
                    $liburDariJadwal[] = $tanggal;
                }
            }

            // Buang tanggal libur (menurut jadwal) dari kategori lain, supaya kode
            // yang salah diisi (mis. 's' pada hari libur) tidak ikut terhitung.
            foreach ($kategoriDariKeterangan as $kategoriKey => $tanggalList) {
                if ($kategoriKey === 'libur') {
                    continue;
                }
                $kategoriDariKeterangan[$kategoriKey] = array_values(array_diff($tanggalList, $liburDariJadwal));
            }

            $liburManual = $kategoriDariKeterangan['libur'] ?? [];
            $libur = array_values(array_unique(array_merge($liburDariJadwal, $liburManual)));
            $this->sortTanggal($libur);

            // tanggal yang SUDAH dijelaskan (selain kode alpa 'a'); libur juga
            // "menjelaskan" ketidakhadiran sehingga tidak dianggap alpa.
            $sudahDijelaskan = $libur;
            foreach ($kategoriDariKeterangan as $kategori => $tanggalList) {
                if (in_array($kategori, ['alpa', 'libur'], true)) {
                    continue;
                }
                $sudahDijelaskan = array_merge($sudahDijelaskan, $tanggalList);
            }

            $alpaOtomatis = array_diff($data['kosong_absen'], $sudahDijelaskan);
            $alpaManual = array_diff($kategoriDariKeterangan['alpa'] ?? [], $liburDariJadwal);
            $alpa = array_values(array_unique(array_merge($alpaOtomatis, $alpaManual)));
            $this->sortTanggal($alpa);

            $rekap[$id] = [
                'id' => $id,
                'nama' => $data['nama'],
                'kategori' => array_merge($kategoriDariKeterangan, ['alpa' => $alpa, 'libur' => $libur]),
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
                $this->sortTanggal($tanggalList);
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
            $this->sortTanggal($tanggalList);
            $bagian[] = 'tgl '.implode(', ', $tanggalList).' '.$teks;
        }

        foreach ($kategori as $kategoriKey => $tanggalList) {
            if (! str_starts_with($kategoriKey, 'lainnya_') || empty($tanggalList)) {
                continue;
            }
            $this->sortTanggal($tanggalList);
            $kode = substr($kategoriKey, strlen('lainnya_'));
            $bagian[] = 'tgl '.implode(', ', $tanggalList)." kode '{$kode}' (kode tidak dikenal, cek CODE_MAP)";
        }

        if (! empty($tidakFingerPagi)) {
            $this->sortTanggal($tidakFingerPagi);
            $bagian[] = 'tgl '.implode(', ', $tidakFingerPagi).' tidak absen pagi';
        }

        if (! empty($tidakFingerSore)) {
            $this->sortTanggal($tidakFingerSore);
            $bagian[] = 'tgl '.implode(', ', $tidakFingerSore).' tidak absen sore';
        }

        return implode('; ', $bagian);
    }

    /**
     * Deteksi bulan dan tahun awal periode dari teks info di sheet.
     * Biasanya baris tertentu mengandung teks seperti "2026-05-26 ~ 2026-06-25"
     * atau "Waktu Ab 2026-05-26 ~ 2026-06-25".
     *
     * @return array{bulan: int, tahun: int}|null
     */
    private function deteksiBulanAwal(Worksheet $sheet, int $scanRow): ?array
    {
        // Scan beberapa baris dan kolom terdekat untuk menemukan pola tanggal
        for ($row = max(1, $scanRow - 2); $row <= $scanRow + 2; $row++) {
            for ($col = 1; $col <= 10; $col++) {
                $colLetter = Coordinate::stringFromColumnIndex($col);
                $value = trim((string) $sheet->getCell($colLetter.$row)->getValue());

                // Cari pola YYYY-MM-DD
                if (preg_match('/(\d{4})-(\d{2})-(\d{2})/', $value, $m)) {
                    return [
                        'bulan' => (int) $m[2],
                        'tahun' => (int) $m[1],
                    ];
                }
            }
        }

        return null;
    }

    /**
     * Cari kolom-kolom yang berisi header tanggal (angka 1-31) pada suatu
     * baris, mulai dari kolom tertentu, sampai kolom kosong berturut-turut.
     *
     * Sekarang mengembalikan key format "d/m" (misal "26/5", "1/6") agar
     * tanggal lintas bulan bisa diurutkan secara kronologis.
     *
     * @param  array{bulan: int, tahun: int}|null  $bulanAwal  info bulan/tahun awal dari deteksiBulanAwal()
     * @return array<string, string> [tanggal_key "d/m" => kolom_excel]
     */
    private function bacaHeaderTanggal(Worksheet $sheet, int $headerRow, int $startColumnIndex, ?array $bulanAwal): array
    {
        $hasil = [];
        $kolom = $startColumnIndex;
        $kosongBerturut = 0;

        $bulanSekarang = $bulanAwal ? $bulanAwal['bulan'] : null;
        $tahunSekarang = $bulanAwal ? $bulanAwal['tahun'] : null;
        $tanggalSebelumnya = null;

        while ($kosongBerturut < 3) {
            $colLetter = Coordinate::stringFromColumnIndex($kolom);
            $value = $sheet->getCell($colLetter.$headerRow)->getValue();

            if (is_numeric($value) && (int) $value >= 1 && (int) $value <= 31) {
                $tanggal = (int) $value;

                // Deteksi pergantian bulan: jika tanggal sekarang < tanggal sebelumnya
                // (misal 31 -> 1), berarti bulan naik
                if ($tanggalSebelumnya !== null && $tanggal < $tanggalSebelumnya && $bulanSekarang !== null) {
                    $bulanSekarang++;
                    if ($bulanSekarang > 12) {
                        $bulanSekarang = 1;
                        $tahunSekarang++;
                    }
                }
                $tanggalSebelumnya = $tanggal;

                // Bangun key: jika bulan diketahui, format "d/m", jika tidak fallback ke angka saja
                if ($bulanSekarang !== null) {
                    $key = $tanggal.'/'.$bulanSekarang;
                } else {
                    $key = (string) $tanggal;
                }

                $hasil[$key] = $colLetter;
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

    /**
     * Sort array tanggal format "d/m" secara kronologis.
     * Menangani lintas bulan dengan benar (misal 30/5 sebelum 1/6).
     *
     * @param  string[]  $tanggalList  array tanggal (dimodifikasi in-place by reference)
     */
    private function sortTanggal(array &$tanggalList): void
    {
        usort($tanggalList, function (string $a, string $b): int {
            $pa = $this->parseTanggalKey($a);
            $pb = $this->parseTanggalKey($b);

            // Bandingkan bulan dulu, lalu tanggal
            if ($pa['bulan'] !== $pb['bulan']) {
                return $pa['bulan'] - $pb['bulan'];
            }

            return $pa['tanggal'] - $pb['tanggal'];
        });
    }

    /**
     * Parse key tanggal "d/m" menjadi array [tanggal, bulan].
     * Fallback: jika format lama (angka saja), bulan = 0.
     *
     * @return array{tanggal: int, bulan: int}
     */
    private function parseTanggalKey(string $key): array
    {
        if (str_contains($key, '/')) {
            $parts = explode('/', $key);

            return [
                'tanggal' => (int) $parts[0],
                'bulan' => (int) ($parts[1] ?? 0),
            ];
        }

        return [
            'tanggal' => (int) $key,
            'bulan' => 0,
        ];
    }
}
