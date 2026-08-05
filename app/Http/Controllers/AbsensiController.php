<?php

namespace App\Http\Controllers;

use App\Services\AbsensiRekapService;
use Illuminate\Http\Request;
use Illuminate\Support\Facades\Log;
use Illuminate\Support\Str;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class AbsensiController extends Controller
{
    public function index()
    {
        return view('absensi.index');
    }

    public function proses(Request $request, AbsensiRekapService $service)
    {
        $request->validate([
            'standard_report' => ['required', 'file', 'mimes:xls,xlsx'],
            'keterangan_lain' => ['required', 'file', 'mimes:xls,xlsx'],
            'template_gaji' => ['required', 'file', 'mimes:xls,xlsx'],
        ], [], [
            'standard_report' => 'file Standard Report',
            'keterangan_lain' => 'file Keterangan Lain',
            'template_gaji' => 'file Template Gaji',
        ]);

        $dir = storage_path('app/tmp-absensi/'.Str::uuid());

        if (! is_dir($dir)) {
            mkdir($dir, 0775, true);
        }

        $standardReportPath = $request->file('standard_report')->move($dir, 'standard_report.'.$request->file('standard_report')->extension());
        $keteranganLainPath = $request->file('keterangan_lain')->move($dir, 'keterangan_lain.'.$request->file('keterangan_lain')->extension());
        $templatePath = $request->file('template_gaji')->move($dir, 'template_gaji.'.$request->file('template_gaji')->extension());

        $outputPath = $dir.'/rekap_absensi.xlsx';

        try {
            $spreadsheet = $service->generate(
                $standardReportPath->getPathname(),
                $keteranganLainPath->getPathname(),
                $templatePath->getPathname()
            );

            (new Xlsx($spreadsheet))->save($outputPath);

            $filename = 'Rekap Absensi - '.now()->format('d-m-Y').'.xlsx';

            return response()->download(
                $outputPath,
                $filename,
                [
                    'X-Download-Filename' => $filename,
                ]
            )->deleteFileAfterSend(true);
        } catch (\Throwable $e) {
            Log::error('AbsensiController: gagal memproses rekap absensi', [
                'error' => $e->getMessage(),
                'trace' => $e->getTraceAsString(),
            ]);

            // Jika request AJAX, kembalikan JSON error
            if ($request->ajax() || $request->wantsJson()) {
                return response()->json([
                    'error' => 'Gagal memproses file. Pastikan format dan isi file sesuai, lalu coba kembali.',
                ], 422);
            }

            return back()
                ->withInput()
                ->with('error', 'Gagal memproses file. Pastikan format dan isi file sesuai, lalu coba kembali.');
        } finally {
            // Hanya hapus file INPUT sementara. File output ($outputPath) TIDAK dihapus di sini
            // karena deleteFileAfterSend(true) sudah mengurusnya setelah response terkirim.
            // Menghapus $outputPath di sini menyebabkan error "file does not exist" saat streaming.
            @unlink($standardReportPath->getPathname());
            @unlink($keteranganLainPath->getPathname());
            @unlink($templatePath->getPathname());
            @rmdir($dir);
        }
    }
}
