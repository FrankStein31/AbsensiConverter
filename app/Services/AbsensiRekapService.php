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
 * DAFTAR KARYAWAN YANG DIREKAP (patokan: Lap. Log Absen):
 * - Sumber kebenaran daftar karyawan adalah sheet "Lap. Log Absen" pada
 *   StandardReport. Keterangan_Lain HANYA berfungsi sebagai data pendukung
 *   (sakit/cuti/izin/libur) untuk ID yang SUDAH ada di Lap. Log Absen.
 * - Jika sebuah ID muncul di Keterangan_Lain tapi TIDAK ditemukan di
 *   Lap. Log Absen, ID tsb DI-SKIP sepenuhnya (tidak ikut direkap maupun
 *   ditulis ke template_gaji) — kemungkinan besar human error (typo ID)
 *   saat mengisi Keterangan_Lain, jadi tidak boleh memunculkan karyawan
 *   baru yang sebenarnya tidak tercatat di mesin fingerprint.
 * - Pencocokan datanya sendiri (StandardReport <-> Keterangan_Lain <->
 *   template_gaji) tetap berdasarkan kecocokan ID, bukan nama/urutan.
 *
 * PENCOCOKAN BARIS (berdasarkan ID, bukan urutan/nama):
 * - template_gaji.xlsx punya kolom ID di kolom AC. Baris existing di
 *   template_gaji TIDAK PERNAH ditimpa No/Nama/ID-nya — hanya kolom
 *   absensi (K,M,O,Q,S,U,W,Y) di baris tsb yang diisi/diperbarui, dicari
 *   berdasarkan kecocokan ID di kolom AC dengan ID dari StandardReport.
 * - Jika ID dari StandardReport TIDAK ditemukan di kolom AC template_gaji
 *   (karyawan baru), baris baru ditambahkan tepat di bawah baris data
 *   terakhir yang sudah terisi di template_gaji (No lanjut, Nama & ID
 *   diisi, lalu kolom absensi).
 * - Baris di template_gaji yang punya Nama tapi kolom ID-nya kosong akan
 *   DILEWATI dari pencocokan (tidak diisi/disentuh sama sekali) karena
 *   patokan pencocokan adalah ID, bukan nama.
 * - Karyawan dari StandardReport yang ID-nya kosong juga dilewati (tidak
 *   bisa dicocokkan maupun ditambahkan tanpa ID).
 *
 * ATURAN LIBUR:
 * - TIDAK lagi memakai sheet "Jadwal Info" — StandardReport hanya dibaca dari
 *   sheet "Lap. Log Absen" saja (untuk ID, dan pola jam masuk/pulang).
 * - Libur ditentukan otomatis oleh KALENDER (library irfa/php-hari-libur):
 *   -- setiap hari Minggu -> libur;
 *   -- setiap tanggal yang menurut library adalah libur nasional / cuti
 *      bersama -> libur.
 *   Aturan ini berlaku sama untuk SEMUA karyawan (bukan per-jadwal individu).
 * - Keterangan_Lain (kode 'l') berfungsi sebagai PENDUKUNG: jika ada tanggal
 *   yang menurut kalender seharusnya hari aktif (bukan Minggu, bukan libur
 *   nasional menurut library) tapi ditandai 'l' secara manual, tanggal itu
 *   TETAP dihitung libur — dipakai untuk cuti bersama/libur nasional lain
 *   yang belum/tidak tercakup di library.
 * - Tanggal yang sudah libur (dari kalender) akan MENIMPA kode apapun
 *   (s/sn/c/i/is) yang mungkin salah diisi di Keterangan_Lain untuk tanggal
 *   tsb.
 * - TAMPILAN kolom "Tanggal Libur" di template TIDAK menyertakan hari
 *   Minggu — hari Minggu tetap dihitung libur secara internal (supaya tidak
 *   dianggap alpa) tapi tidak dicantumkan satu-satu di kolom. Yang tampil di
 *   kolom hanya libur pada hari aktif Senin-Sabtu (libur nasional/cuti
 *   bersama, termasuk yang ditandai manual kode 'l').
 *
 * FORMAT TANGGAL:
 * - Semua tanggal disimpan secara internal sebagai "d/m" (contoh: "26/5",
 *   "1/6") agar jelas bulannya dan bisa diurutkan kronologis lintas bulan.
 * - Untuk DITAMPILKAN (kolom tanggal di template_gaji & Catatan), daftar
 *   tanggal ini dirapikan dan dikelompokkan per bulan dengan nama bulan
 *   Indonesia, contoh: "26/5,28/5,29/5,1/6,2/6" -> "26, 28, 29 Mei, 1, 2
 *   Juni". Lihat formatTanggalUntukTampilan(). Jumlah koma di hasil akhir
 *   tetap sama dengan jumlah tanggal dikurangi 1, sehingga formula hitung
 *   jumlah hari di template (mis. kolom L/N/P) tetap berfungsi benar.
 */
class AbsensiRekapService
{
    /** Nama bulan Indonesia untuk formatTanggalUntukTampilan(). */
    private const NAMA_BULAN = [
        1 => 'Januari', 2 => 'Februari', 3 => 'Maret', 4 => 'April',
        5 => 'Mei', 6 => 'Juni', 7 => 'Juli', 8 => 'Agustus',
        9 => 'September', 10 => 'Oktober', 11 => 'November', 12 => 'Desember',
    ];

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

    /** Kolom berisi ID karyawan di template_gaji.xlsx — patokan pencocokan baris. */
    private const TEMPLATE_ID_COLUMN = 'AC';

    public function generate(string $standardReportPath, string $keteranganLainPath, string $templatePath): Spreadsheet
    {
        $standardReport = $this->parseStandardReport($standardReportPath);
        $keterangan = $this->parseKeteranganLain($keteranganLainPath, $standardReport['bulan_awal']);
        $rekap = $this->gabungkan(
            $standardReport['karyawan'],
            $standardReport['libur_global'],
            $standardReport['libur_minggu'],
            $keterangan
        );

        return $this->isiTemplate($templatePath, $rekap);
    }

    /**
     * @return array{
     *     karyawan: array<string, array{id:string, nama:string, kosong_absen: string[], tidak_finger_pagi: string[], tidak_finger_sore: string[]}>,
     *     libur_global: string[],
     *     libur_minggu: string[],
     *     bulan_awal: array{bulan: int, tahun: int}|null
     * }
     *                keyed by ID karyawan. "kosong_absen" = tanggal (di luar libur global) yang tidak
     *                ada satu pun log jam masuk/pulang. "libur_global" = daftar tanggal ("d/m") yang
     *                menurut kalender adalah hari Minggu atau libur nasional/cuti bersama (dipakai
     *                untuk logika internal). "libur_minggu" = subset khusus hari Minggu saja (dipakai
     *                untuk menyaring kolom "Tanggal Libur" agar Minggu tidak ditampilkan di sana).
     */
    private function parseStandardReport(string $path): array
    {
        $spreadsheet = IOFactory::load($path);

        $logSheet = $spreadsheet->getSheetByName('Lap. Log Absen');

        if (! $logSheet) {
            throw new RuntimeException(
                'Format 1_StandardReport tidak dikenali: sheet "Lap. Log Absen" tidak ditemukan.'
            );
        }

        // --- Ambil bulan awal dari info periode di sheet (biasanya baris 3, misal "2026-05-26 ~ 2026-06-25")
        $bulanAwalLog = $this->deteksiBulanAwal($logSheet, 3);

        // --- Sheet "Lap. Log Absen": blok per-karyawan (baris "ID:" lalu baris jam di bawahnya).
        $tanggalLog = $this->bacaHeaderTanggal($logSheet, 4, 1, $bulanAwalLog);

        // --- Libur global (Minggu + libur nasional/cuti bersama dari kalender), berlaku untuk semua karyawan.
        $liburHasil = $this->hitungLiburGlobal($tanggalLog);
        $liburGlobal = $liburHasil['semua'];
        $liburMinggu = $liburHasil['minggu'];
        $liburGlobalSet = array_flip($liburGlobal);

        $karyawan = [];
        $highestRow = $logSheet->getHighestRow();

        for ($r = 1; $r <= $highestRow; $r++) {
            if (trim((string) $logSheet->getCell('A'.$r)->getValue()) !== 'ID:') {
                continue;
            }
            $id = trim((string) $logSheet->getCell('C'.$r)->getValue());
            if ($id === '') {
                continue;
            }
            // Nama dicari fleksibel di baris blok yang sama (layout tiap mesin fingerprint
            // bisa beda kolom), dengan mencari label "Nama"/"Name" lalu ambil sel sesudahnya.
            $nama = $this->cariNilaiSetelahLabel($logSheet, $r, ['nama', 'name']) ?? '';
            $dataRow = $r + 1;

            if (! isset($karyawan[$id])) {
                $karyawan[$id] = [
                    'id' => $id,
                    'nama' => $nama,
                    'kosong_absen' => [],
                    'tidak_finger_pagi' => [],
                    'tidak_finger_sore' => [],
                ];
            } elseif ($nama !== '' && $karyawan[$id]['nama'] === '') {
                $karyawan[$id]['nama'] = $nama;
            }

            foreach ($tanggalLog as $tanggalKey => $info) {
                if (isset($liburGlobalSet[$tanggalKey])) {
                    continue; // hari libur (Minggu/libur nasional) -> tidak dicek kehadirannya
                }

                $jam = trim((string) $logSheet->getCell($info['kolom'].$dataRow)->getValue());

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

        return [
            'karyawan' => $karyawan,
            'libur_global' => $liburGlobal,
            'libur_minggu' => $liburMinggu,
            'bulan_awal' => $bulanAwalLog,
        ];
    }

    /**
     * Hitung daftar tanggal ("d/m") yang termasuk libur GLOBAL (berlaku untuk
     * semua karyawan): hari Minggu, atau libur nasional/cuti bersama menurut
     * library kalender Indonesia (irfa/php-hari-libur).
     *
     * Mengembalikan 'semua' (Minggu + libur nasional, dipakai untuk logika
     * internal seperti pengecualian alpa) dan 'minggu' (khusus hari Minggu
     * saja, dipakai untuk MENYARING kolom "Tanggal Libur" di template agar
     * hari Minggu tidak ikut ditampilkan di sana — lihat gabungkan()).
     *
     * @param  array<string, array{kolom:string, tanggal:int, bulan:int, tahun:?int}>  $tanggalInfo
     * @return array{semua: string[], minggu: string[]}
     */
    private function hitungLiburGlobal(array $tanggalInfo): array
    {
        $liburGlobal = [];
        $liburMinggu = [];

        foreach ($tanggalInfo as $tanggalKey => $info) {
            if ($info['tahun'] === null) {
                // Tahun tidak diketahui (info periode tidak terbaca) -> tidak bisa dicek ke
                // kalender. Tetap bisa ditandai libur manual lewat kode 'l' di Keterangan_Lain.
                continue;
            }

            $timestamp = mktime(0, 0, 0, $info['bulan'], $info['tanggal'], $info['tahun']);
            $isMinggu = (int) date('N', $timestamp) === 7;

            // Karena library hari libur (irfa/php-hari-libur) tidak kompatibel dengan PHP 8.5,
            // pendeteksian otomatis libur nasional dilewati.
            // Libur nasional dapat ditandai secara manual dengan kode 'l' pada file Keterangan_Lain.
            $isLiburNasional = false;

            if ($isMinggu) {
                $liburMinggu[] = $tanggalKey;
            }

            if ($isMinggu || $isLiburNasional) {
                $liburGlobal[] = $tanggalKey;
            }
        }

        return [
            'semua' => $liburGlobal,
            'minggu' => $liburMinggu,
        ];
    }

    /**
     * Cari label tertentu (mis. "Nama"/"Name") pada suatu baris, lalu ambil nilai
     * di kolom sesudahnya. Dipakai untuk membaca Nama karyawan pada baris blok
     * "ID:" di sheet "Lap. Log Absen" tanpa bergantung pada kolom tetap, karena
     * layout blok bisa berbeda-beda antar mesin fingerprint.
     *
     * @param  string[]  $labelKeywords  daftar keyword label (huruf kecil, tanpa titik dua) yang dicari
     */
    private function cariNilaiSetelahLabel(Worksheet $sheet, int $row, array $labelKeywords): ?string
    {
        for ($col = 1; $col <= 40; $col++) {
            $colLetter = Coordinate::stringFromColumnIndex($col);
            $value = strtolower(trim((string) $sheet->getCell($colLetter.$row)->getValue()));
            $value = rtrim($value, ':');

            if (in_array($value, $labelKeywords, true)) {
                $nextCol = Coordinate::stringFromColumnIndex($col + 1);
                $nilai = trim((string) $sheet->getCell($nextCol.$row)->getValue());

                if ($nilai !== '') {
                    return $nilai;
                }
            }
        }

        return null;
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
     * @return array<string, array{nama: string, kategori: array<string, string[]>}>
     *                                                                               [id_karyawan] => nama + [kategori] = daftar tanggal (format "d/m")
     */
    private function parseKeteranganLain(string $path, ?array $bulanAwalFallback = null): array
    {
        $spreadsheet = IOFactory::load($path);
        $sheet = $spreadsheet->getSheetByName('SCIL') ?? $spreadsheet->getSheet(0);

        $bulanAwal = $this->deteksiBulanAwal($sheet, 3) ?? $bulanAwalFallback;
        $tanggalCol = $this->bacaHeaderTanggal($sheet, 4, 3, $bulanAwal);

        $hasil = [];
        $row = 5;
        while (true) {
            $id = trim((string) $sheet->getCell('A'.$row)->getValue());
            if ($id === '') {
                break;
            }

            $nama = trim((string) $sheet->getCell('B'.$row)->getValue());
            if (! isset($hasil[$id])) {
                $hasil[$id] = ['nama' => $nama, 'kategori' => []];
            } elseif ($nama !== '' && $hasil[$id]['nama'] === '') {
                $hasil[$id]['nama'] = $nama;
            }

            foreach ($tanggalCol as $tanggalKey => $info) {
                $kodeMentah = strtolower(trim((string) $sheet->getCell($info['kolom'].$row)->getValue()));
                if ($kodeMentah === '') {
                    continue;
                }
                $kategori = self::CODE_MAP[$kodeMentah] ?? null;
                if ($kategori === null) {
                    // kode tidak dikenal -> tetap dicatat sebagai "lainnya" supaya tidak hilang diam-diam
                    $kategori = 'lainnya_'.$kodeMentah;
                }
                $hasil[$id]['kategori'][$kategori][] = $tanggalKey;
            }
            $row++;
        }

        return $hasil;
    }

    /**
     * Gabungkan hasil StandardReport (Lap. Log Absen) + Keterangan_Lain jadi satu rekap final per karyawan.
     *
     * PATOKAN DAFTAR KARYAWAN: Lap. Log Absen (StandardReport), BUKAN Keterangan_Lain.
     * Jika sebuah ID hanya muncul di Keterangan_Lain tapi TIDAK ada di Lap. Log Absen,
     * ID tsb DI-SKIP sepenuhnya (tidak ikut direkap/ditulis ke template) — dianggap
     * kemungkinan human error (typo ID, dsb) di Keterangan_Lain. Pencocokan datanya
     * sendiri tetap berdasarkan ID (bukan nama/urutan), sama seperti sebelumnya.
     *
     * Libur GLOBAL (Minggu + libur nasional/cuti bersama dari kalender) berlaku sama untuk
     * semua karyawan dan menimpa kategori lain di tanggal yang sama apabila ada kode yang
     * salah diisi pada hari libur. Kode 'l' manual di Keterangan_Lain menjadi tambahan/pendukung
     * untuk tanggal yang menurut kalender adalah hari aktif tapi ternyata tetap libur (mis.
     * cuti bersama/libur nasional yang belum tercakup library).
     *
     * Kolom "Tanggal Libur" yang ditulis ke template TIDAK menyertakan hari Minggu (Minggu
     * tetap dihitung libur secara internal supaya tidak dianggap alpa, tapi tidak perlu
     * dicantumkan satu-satu di kolom karena memang libur mingguan tetap) — yang ditampilkan
     * hanya libur yang jatuh di hari aktif Senin-Sabtu (mis. libur nasional/cuti bersama yang
     * ditandai kode 'l').
     *
     * @param  array<string, array{id:string, nama:string, kosong_absen:string[], tidak_finger_pagi:string[], tidak_finger_sore:string[]}>  $karyawanLog  hasil parseStandardReport()['karyawan']
     * @param  string[]  $liburGlobal  hasil parseStandardReport()['libur_global']
     * @param  string[]  $liburMinggu  hasil parseStandardReport()['libur_minggu']
     * @param  array<string, array{nama:string, kategori:array<string,string[]>}>  $keterangan  hasil parseKeteranganLain()
     */
    private function gabungkan(array $karyawanLog, array $liburGlobal, array $liburMinggu, array $keterangan): array
    {
        $rekap = [];

        // Patokan daftar ID karyawan HANYA dari StandardReport (Lap. Log Absen).
        // ID yang hanya ada di Keterangan_Lain (tidak ada di Lap. Log Absen) di-skip.
        $semuaId = array_keys($karyawanLog);

        foreach ($semuaId as $id) {
            $dataLog = $karyawanLog[$id] ?? null;
            $dataKeterangan = $keterangan[$id] ?? null;

            $kategoriDariKeterangan = $dataKeterangan['kategori'] ?? [];

            $nama = $dataLog['nama'] ?? '';
            if ($nama === '') {
                $nama = $dataKeterangan['nama'] ?? '';
            }

            // Buang tanggal libur global dari kategori lain, supaya kode yang salah diisi
            // (mis. 's' pada hari Minggu/libur nasional) tidak ikut terhitung.
            foreach ($kategoriDariKeterangan as $kategoriKey => $tanggalList) {
                if ($kategoriKey === 'libur') {
                    continue;
                }
                $kategoriDariKeterangan[$kategoriKey] = array_values(array_diff($tanggalList, $liburGlobal));
            }

            // Kode 'l' manual di Keterangan_Lain (pendukung untuk cuti bersama/libur nasional
            // lain di luar Minggu & yang sudah dikenali library) digabung dengan libur global.
            $liburManual = $kategoriDariKeterangan['libur'] ?? [];
            $libur = array_values(array_unique(array_merge($liburGlobal, $liburManual)));
            $this->sortTanggal($libur);

            // Untuk KOLOM "Tanggal Libur" yang ditampilkan ke user: buang hari Minggu.
            // Minggu tidak perlu dicantumkan satu-satu karena memang libur mingguan tetap;
            // yang ditampilkan hanya libur di hari aktif Senin-Sabtu.
            $liburTampil = array_values(array_diff($libur, $liburMinggu));
            $this->sortTanggal($liburTampil);

            // tanggal yang SUDAH dijelaskan (selain kode alpa 'a'); libur (termasuk Minggu)
            // tetap "menjelaskan" ketidakhadiran sehingga tidak dianggap alpa, walau Minggu
            // tidak ditampilkan di kolom "Tanggal Libur".
            $sudahDijelaskan = $libur;
            foreach ($kategoriDariKeterangan as $kategori => $tanggalList) {
                if (in_array($kategori, ['alpa', 'libur'], true)) {
                    continue;
                }
                $sudahDijelaskan = array_merge($sudahDijelaskan, $tanggalList);
            }

            $kosongAbsen = $dataLog['kosong_absen'] ?? [];
            $alpaOtomatis = array_diff($kosongAbsen, $sudahDijelaskan);
            $alpaManual = array_diff($kategoriDariKeterangan['alpa'] ?? [], $liburGlobal);
            $alpa = array_values(array_unique(array_merge($alpaOtomatis, $alpaManual)));
            $this->sortTanggal($alpa);

            $rekap[$id] = [
                'id' => $id,
                'nama' => $nama,
                'kategori' => array_merge($kategoriDariKeterangan, ['alpa' => $alpa, 'libur' => $liburTampil]),
                'tidak_finger_pagi' => $dataLog['tidak_finger_pagi'] ?? [],
                'tidak_finger_sore' => $dataLog['tidak_finger_sore'] ?? [],
            ];
        }

        return $rekap;
    }

    private function isiTemplate(string $templatePath, array $rekap): Spreadsheet
    {
        $spreadsheet = IOFactory::load($templatePath);
        $sheet = $spreadsheet->getSheet(0);

        // --- 1. Petakan baris existing di template_gaji.xlsx berdasarkan ID
        //     (kolom AC), sekaligus cari baris kosong berikutnya (untuk
        //     karyawan baru) dan No terakhir (untuk lanjutan penomoran).
        $idKeRow = [];
        $lastUsedRow = self::FIRST_DATA_ROW - 1;
        $lastNo = 0;

        $row = self::FIRST_DATA_ROW;
        while (true) {
            $nama = trim((string) $sheet->getCell('B'.$row)->getValue());
            $idExisting = trim((string) $sheet->getCell(self::TEMPLATE_ID_COLUMN.$row)->getValue());

            if ($nama === '' && $idExisting === '') {
                break; // baris kosong pertama = akhir data existing di template
            }

            $lastUsedRow = $row;

            $noValue = $sheet->getCell('A'.$row)->getValue();
            if (is_numeric($noValue)) {
                $lastNo = max($lastNo, (int) $noValue);
            }

            if ($idExisting !== '') {
                $idKeRow[$idExisting] = $row;
            }
            // Baris dengan Nama tapi ID kosong SENGAJA tidak dimasukkan ke
            // $idKeRow — dilewati dari pencocokan karena patokannya ID.

            $row++;
        }

        $nextNewRow = $lastUsedRow + 1;
        $nextNo = $lastNo + 1;

        // --- 2. Isi kolom absensi per karyawan, dicocokkan berdasarkan ID.
        foreach ($rekap as $data) {
            $id = trim((string) $data['id']);
            if ($id === '') {
                continue; // tanpa ID tidak bisa dicocokkan/ditambahkan, skip.
            }

            if (isset($idKeRow[$id])) {
                // ID sudah ada di template: HANYA isi kolom absensi di baris
                // itu. No, Nama, ID, dan kolom identitas lain (C-G) yang
                // sudah ada di template TIDAK disentuh/ditimpa.
                $targetRow = $idKeRow[$id];
            } else {
                // ID baru (tidak ada di template): tambahkan baris baru tepat
                // di bawah data terakhir, isi No, Nama, dan ID-nya.
                $targetRow = $nextNewRow;
                $sheet->setCellValue('A'.$targetRow, $nextNo);
                $sheet->setCellValue('B'.$targetRow, $data['nama']);
                $sheet->setCellValue(self::TEMPLATE_ID_COLUMN.$targetRow, $id);
                // C, D, E, F, G (No Rekening, Jabatan, Area, Gaji Pokok, Bonus) SENGAJA dikosongkan.

                $idKeRow[$id] = $targetRow; // jaga-jaga bila ID sama muncul dobel di rekap
                $nextNewRow++;
                $nextNo++;
            }

            foreach (self::TEMPLATE_DATE_COLUMN as $kategori => $kolom) {
                $tanggalList = $data['kategori'][$kategori] ?? [];
                $this->sortTanggal($tanggalList);
                $sheet->setCellValue($kolom.$targetRow, $this->formatTanggalUntukTampilan($tanggalList));
            }

            $sheet->setCellValue('Y'.$targetRow, $this->buatCatatan(
                $data['kategori'],
                $data['tidak_finger_pagi'],
                $data['tidak_finger_sore']
            ));
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
            $bagian[] = 'tgl '.$this->formatTanggalUntukTampilan($tanggalList).' '.$teks;
        }

        foreach ($kategori as $kategoriKey => $tanggalList) {
            if (! str_starts_with($kategoriKey, 'lainnya_') || empty($tanggalList)) {
                continue;
            }
            $this->sortTanggal($tanggalList);
            $kode = substr($kategoriKey, strlen('lainnya_'));
            $bagian[] = 'tgl '.$this->formatTanggalUntukTampilan($tanggalList)." kode '{$kode}' (kode tidak dikenal, cek CODE_MAP)";
        }

        if (! empty($tidakFingerPagi)) {
            $this->sortTanggal($tidakFingerPagi);
            $bagian[] = 'tgl '.$this->formatTanggalUntukTampilan($tidakFingerPagi).' tidak absen pagi';
        }

        if (! empty($tidakFingerSore)) {
            $this->sortTanggal($tidakFingerSore);
            $bagian[] = 'tgl '.$this->formatTanggalUntukTampilan($tidakFingerSore).' tidak absen sore';
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
     * Mengembalikan key format "d/m" (misal "26/5", "1/6") agar tanggal
     * lintas bulan bisa diurutkan secara kronologis, beserta info tanggal
     * lengkap (kolom Excel, tanggal, bulan, tahun) supaya bisa dicocokkan
     * ke kalender (mis. cek hari Minggu / libur nasional).
     *
     * @param  array{bulan: int, tahun: int}|null  $bulanAwal  info bulan/tahun awal dari deteksiBulanAwal()
     * @return array<string, array{kolom: string, tanggal: int, bulan: ?int, tahun: ?int}>
     *                                                                                     [tanggal_key "d/m" => info]
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

                $hasil[$key] = [
                    'kolom' => $colLetter,
                    'tanggal' => $tanggal,
                    'bulan' => $bulanSekarang,
                    'tahun' => $tahunSekarang,
                ];
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
     * Format daftar tanggal (format internal "d/m") menjadi teks yang lebih
     * mudah dibaca, dikelompokkan per bulan dengan nama bulan Indonesia.
     *
     * Contoh: ["26/5","28/5","29/5","1/6","2/6","3/6","4/6","5/6"]
     *      -> "26, 28, 29 Mei, 1, 2, 3, 4, 5 Juni"
     *
     * Catatan: jumlah koma pada hasil akhir tetap sama dengan (jumlah
     * tanggal - 1), sama seperti format lama "d/m,d/m,...", sehingga
     * formula hitung jumlah hari di template_gaji.xlsx (yang menghitung
     * koma) tetap menghasilkan angka yang benar.
     *
     * @param  string[]  $tanggalList  idealnya sudah diurutkan lewat sortTanggal()
     */
    private function formatTanggalUntukTampilan(array $tanggalList): string
    {
        if (empty($tanggalList)) {
            return '';
        }

        $kelompokPerBulan = []; // bulan (int) => daftar hari (int)
        $urutanBulan = [];

        foreach ($tanggalList as $tanggalKey) {
            $parsed = $this->parseTanggalKey($tanggalKey);
            $bulan = $parsed['bulan'];

            if (! isset($kelompokPerBulan[$bulan])) {
                $kelompokPerBulan[$bulan] = [];
                $urutanBulan[] = $bulan;
            }

            $kelompokPerBulan[$bulan][] = $parsed['tanggal'];
        }

        $bagian = [];
        foreach ($urutanBulan as $bulan) {
            $namaBulan = self::NAMA_BULAN[$bulan] ?? (string) $bulan;
            $bagian[] = implode(', ', $kelompokPerBulan[$bulan]).' '.$namaBulan;
        }

        return implode(', ', $bagian);
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