<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta name="robots" content="noindex, nofollow">
    <title>Konverter Rekap Absensi</title>
    @if (file_exists(public_path('build/manifest.json')) || file_exists(public_path('hot')))
        @vite(['resources/css/app.css', 'resources/js/app.js'])
    @else
        <script src="https://unpkg.com/@tailwindcss/browser@4"></script>
    @endif
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap');
        
        body { 
            font-family: 'Plus Jakarta Sans', ui-sans-serif, system-ui, sans-serif; 
            background-color: #f8fafc;
            background-image: radial-gradient(#e2e8f0 1px, transparent 1px);
            background-size: 24px 24px;
        }

        /* Drop zone states - Enhanced with softer colors */
        .drop-zone { transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1); }
        .drop-zone:hover { border-color: #94a3b8; background-color: #f8fafc; }
        
        .drop-zone.drag-over {
            border-color: #6366f1 !important; /* Indigo for drag over */
            background-color: #eef2ff !important;
            transform: scale(1.02);
        }
        .drop-zone.drag-over .drop-icon-card {
            transform: translate(0, -6px) scale(1.05);
            box-shadow: 0 10px 25px -5px rgba(99, 102, 241, 0.2);
            border-color: #c7d2fe;
        }
        .drop-zone.drag-over .drop-icon-card svg {
            color: #6366f1;
        }
        .drop-zone.drag-over .drop-icon-card-ghost {
            opacity: 1 !important;
            transform: scale(1.1);
        }
        
        .drop-zone.has-file {
            border-color: #10b981 !important; /* Emerald for success */
            background-color: #f0fdf4 !important;
        }
        .drop-zone.has-file .drop-icon-card {
            background-color: #d1fae5 !important;
            border-color: #a7f3d0;
            box-shadow: none;
        }
        .drop-zone.has-file svg {
            color: #059669 !important;
        }

        /* Icon card animation */
        .drop-icon-card { 
            transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
            border: 1px solid #e2e8f0; 
        }
        .drop-icon-card-ghost { transition: all 0.3s ease; }

        /* Loading spinner */
        @keyframes spin { to { transform: rotate(360deg); } }
        .spinner {
            display: inline-block;
            width: 16px; height: 16px;
            border: 2px solid rgba(255,255,255,0.3);
            border-top-color: white;
            border-radius: 50%;
            animation: spin 0.6s linear infinite;
            vertical-align: middle;
            margin-left: 8px;
        }

        /* File chip fade-in */
        @keyframes fadeSlideIn {
            from { opacity: 0; transform: translateY(8px); }
            to   { opacity: 1; transform: translateY(0); }
        }
        .file-chip { animation: fadeSlideIn 0.3s cubic-bezier(0.4, 0, 0.2, 1) forwards; }
    </style>
</head>
<body class="min-h-screen antialiased text-slate-600 relative overflow-x-hidden">

    {{-- Top Decorative Glow --}}
    <div class="absolute top-0 inset-x-0 h-40 bg-gradient-to-b from-indigo-50/80 to-transparent -z-10"></div>

    {{-- Header --}}
    <header class="sticky top-0 z-40 bg-white/80 backdrop-blur-xl border-b border-slate-200/80 px-6 py-3.5 shadow-sm">
        <div class="max-w-4xl mx-auto flex items-center justify-between">
            <div class="flex items-center gap-3">
                <div class="w-8 h-8 bg-gradient-to-br from-slate-800 to-slate-900 rounded-lg flex items-center justify-center shrink-0 shadow-md shadow-slate-900/10">
                    <svg class="w-4 h-4 text-white" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24">
                        <path stroke-linecap="round" stroke-linejoin="round" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"/>
                    </svg>
                </div>
                <span class="text-sm font-bold text-slate-800 tracking-tight">Konverter Absensi</span>
            </div>
        </div>
    </header>

    <main class="max-w-4xl mx-auto px-4 py-10 sm:py-14">

        {{-- Main Card Container --}}
        <div class="bg-white rounded-2xl shadow-[0_8px_30px_rgb(0,0,0,0.04)] border border-slate-100 p-6 sm:p-8">
            
            {{-- Page Heading --}}
            <div class="mb-8 border-b border-slate-100 pb-6">
                <h1 class="text-xl sm:text-2xl font-bold text-slate-900 tracking-tight">Generate Rekap Absensi</h1>
                <p class="mt-2 text-sm text-slate-500 leading-relaxed max-w-2xl">
                    Unggah 3 file Excel pendukung: <strong>Standard Report</strong> dari mesin, <strong>Keterangan Lain</strong> (sakit/izin), dan <strong>Template Gaji</strong>. Sistem akan menggabungkannya menjadi satu file rekap siap unduh.
                </p>
            </div>

            {{-- Error / Validation Alerts --}}
            @if (session('error'))
                <div class="mb-6 flex items-start gap-3 rounded-xl border border-red-100 bg-red-50/80 px-4 py-3 animate-[fadeSlideIn_0.3s_ease-out]">
                    <div class="bg-red-100 p-1 rounded-full shrink-0 mt-0.5">
                        <svg class="w-4 h-4 text-red-600" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24">
                            <path stroke-linecap="round" stroke-linejoin="round" d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-3L13.732 4c-.77-1.333-2.694-1.333-3.464 0L3.34 16c-.77 1.333.192 3 1.732 3z"/>
                        </svg>
                    </div>
                    <div>
                        <p class="text-sm font-bold text-red-800">Gagal Memproses</p>
                        <p class="text-xs text-red-600 mt-1">{{ session('error') }}</p>
                    </div>
                </div>
            @endif

            @if ($errors->any())
                <div class="mb-6 rounded-xl border border-red-100 bg-red-50/80 px-4 py-3 animate-[fadeSlideIn_0.3s_ease-out]">
                    <p class="text-sm font-bold text-red-800 mb-2">Periksa input berikut:</p>
                    <ul class="space-y-1.5">
                        @foreach ($errors->all() as $error)
                            <li class="text-xs text-red-600 flex items-center gap-2">
                                <span class="w-1.5 h-1.5 bg-red-400 rounded-full shrink-0"></span>{{ $error }}
                            </li>
                        @endforeach
                    </ul>
                </div>
            @endif

            <form
                method="POST"
                action="{{ route('absensi.proses') }}"
                enctype="multipart/form-data"
                id="form-absensi"
                novalidate
            >
                @csrf

                <div class="grid grid-cols-1 md:grid-cols-3 gap-6 mb-8">

                    {{-- Drop Zone 1: Standard Report --}}
                    <div>
                        <div class="flex items-center gap-2 mb-2">
                            <span class="flex items-center justify-center w-5 h-5 bg-indigo-100 text-indigo-700 rounded-full text-[10px] font-bold shrink-0">1</span>
                            <p class="text-sm font-semibold text-slate-800">Standard Report</p>
                        </div>
                        <div
                            class="drop-zone relative rounded-xl border-2 border-dashed border-slate-200 bg-slate-50/50 cursor-pointer overflow-hidden group"
                            data-target="standard_report"
                            id="zone-standard_report"
                        >
                            <div class="relative z-10 flex flex-col items-center justify-center py-8 px-4 text-center min-h-[160px]">
                                <div class="relative w-14 h-14 mb-3">
                                    <div class="drop-icon-card absolute inset-0 bg-white rounded-xl shadow-sm flex items-center justify-center z-10">
                                        <svg class="w-6 h-6 text-slate-400 group-hover:text-indigo-500 transition-colors" fill="none" stroke="currentColor" stroke-width="1.5" viewBox="0 0 24 24">
                                            <path stroke-linecap="round" stroke-linejoin="round" d="M19.5 14.25v-2.625a3.375 3.375 0 00-3.375-3.375h-1.5A1.125 1.125 0 0113.5 7.125v-1.5a3.375 3.375 0 00-3.375-3.375H8.25m2.25 0H5.625c-.621 0-1.125.504-1.125 1.125v17.25c0 .621.504 1.125 1.125 1.125h12.75c.621 0 1.125-.504 1.125-1.125V11.25a9 9 0 00-9-9z"/>
                                        </svg>
                                    </div>
                                    <div class="drop-icon-card-ghost absolute inset-0 rounded-xl border-2 border-dashed border-indigo-400 opacity-0 z-0"></div>
                                </div>
                                <p class="text-sm font-semibold text-slate-700" id="label-standard_report">Tarik file ke sini</p>
                                <p class="text-[11px] text-slate-500 mt-1">atau klik untuk menelusuri</p>
                                <div class="mt-3 px-2 py-0.5 rounded text-[10px] font-medium bg-slate-200/70 text-slate-500">.xls / .xlsx</div>
                            </div>
                            <input type="file" name="standard_report" id="standard_report" accept=".xls,.xlsx" class="hidden">
                        </div>
                        @error('standard_report')
                            <p class="mt-2 text-xs font-medium text-red-500 flex items-center gap-1"><svg class="w-3 h-3" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 8v4m0 4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z" /></svg>{{ $message }}</p>
                        @enderror
                        <div id="chip-standard_report" class="hidden mt-2 file-chip flex items-center gap-2 bg-indigo-50 border border-indigo-100 rounded-lg px-3 py-2">
                            <svg class="w-4 h-4 text-indigo-500 shrink-0" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24">
                                <path stroke-linecap="round" stroke-linejoin="round" d="M9 12.75L11.25 15 15 9.75M21 12a9 9 0 11-18 0 9 9 0 0118 0z"/>
                            </svg>
                            <span class="text-xs font-medium text-indigo-700 truncate" id="chipname-standard_report"></span>
                        </div>
                    </div>

                    {{-- Drop Zone 2: Keterangan Lain --}}
                    <div>
                        <div class="flex items-center gap-2 mb-2">
                            <span class="flex items-center justify-center w-5 h-5 bg-indigo-100 text-indigo-700 rounded-full text-[10px] font-bold shrink-0">2</span>
                            <p class="text-sm font-semibold text-slate-800">Keterangan Lain</p>
                        </div>
                        <div
                            class="drop-zone relative rounded-xl border-2 border-dashed border-slate-200 bg-slate-50/50 cursor-pointer overflow-hidden group"
                            data-target="keterangan_lain"
                            id="zone-keterangan_lain"
                        >
                            <div class="relative z-10 flex flex-col items-center justify-center py-8 px-4 text-center min-h-[160px]">
                                <div class="relative w-14 h-14 mb-3">
                                    <div class="drop-icon-card absolute inset-0 bg-white rounded-xl shadow-sm flex items-center justify-center z-10">
                                        <svg class="w-6 h-6 text-slate-400 group-hover:text-indigo-500 transition-colors" fill="none" stroke="currentColor" stroke-width="1.5" viewBox="0 0 24 24">
                                            <path stroke-linecap="round" stroke-linejoin="round" d="M19.5 14.25v-2.625a3.375 3.375 0 00-3.375-3.375h-1.5A1.125 1.125 0 0113.5 7.125v-1.5a3.375 3.375 0 00-3.375-3.375H8.25m2.25 0H5.625c-.621 0-1.125.504-1.125 1.125v17.25c0 .621.504 1.125 1.125 1.125h12.75c.621 0 1.125-.504 1.125-1.125V11.25a9 9 0 00-9-9z"/>
                                        </svg>
                                    </div>
                                    <div class="drop-icon-card-ghost absolute inset-0 rounded-xl border-2 border-dashed border-indigo-400 opacity-0 z-0"></div>
                                </div>
                                <p class="text-sm font-semibold text-slate-700" id="label-keterangan_lain">Tarik file ke sini</p>
                                <p class="text-[11px] text-slate-500 mt-1">atau klik untuk menelusuri</p>
                                <div class="mt-3 px-2 py-0.5 rounded text-[10px] font-medium bg-slate-200/70 text-slate-500">.xls / .xlsx</div>
                            </div>
                            <input type="file" name="keterangan_lain" id="keterangan_lain" accept=".xls,.xlsx" class="hidden">
                        </div>
                        @error('keterangan_lain')
                            <p class="mt-2 text-xs font-medium text-red-500 flex items-center gap-1"><svg class="w-3 h-3" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 8v4m0 4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z" /></svg>{{ $message }}</p>
                        @enderror
                        <div id="chip-keterangan_lain" class="hidden mt-2 file-chip flex items-center gap-2 bg-indigo-50 border border-indigo-100 rounded-lg px-3 py-2">
                            <svg class="w-4 h-4 text-indigo-500 shrink-0" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24">
                                <path stroke-linecap="round" stroke-linejoin="round" d="M9 12.75L11.25 15 15 9.75M21 12a9 9 0 11-18 0 9 9 0 0118 0z"/>
                            </svg>
                            <span class="text-xs font-medium text-indigo-700 truncate" id="chipname-keterangan_lain"></span>
                        </div>
                    </div>

                    {{-- Drop Zone 3: Template Gaji --}}
                    <div>
                        <div class="flex items-center gap-2 mb-2">
                            <span class="flex items-center justify-center w-5 h-5 bg-indigo-100 text-indigo-700 rounded-full text-[10px] font-bold shrink-0">3</span>
                            <p class="text-sm font-semibold text-slate-800">Template Gaji</p>
                        </div>
                        <div
                            class="drop-zone relative rounded-xl border-2 border-dashed border-slate-200 bg-slate-50/50 cursor-pointer overflow-hidden group"
                            data-target="template_gaji"
                            id="zone-template_gaji"
                        >
                            <div class="relative z-10 flex flex-col items-center justify-center py-8 px-4 text-center min-h-[160px]">
                                <div class="relative w-14 h-14 mb-3">
                                    <div class="drop-icon-card absolute inset-0 bg-white rounded-xl shadow-sm flex items-center justify-center z-10">
                                        <svg class="w-6 h-6 text-slate-400 group-hover:text-indigo-500 transition-colors" fill="none" stroke="currentColor" stroke-width="1.5" viewBox="0 0 24 24">
                                            <path stroke-linecap="round" stroke-linejoin="round" d="M19.5 14.25v-2.625a3.375 3.375 0 00-3.375-3.375h-1.5A1.125 1.125 0 0113.5 7.125v-1.5a3.375 3.375 0 00-3.375-3.375H8.25m2.25 0H5.625c-.621 0-1.125.504-1.125 1.125v17.25c0 .621.504 1.125 1.125 1.125h12.75c.621 0 1.125-.504 1.125-1.125V11.25a9 9 0 00-9-9z"/>
                                        </svg>
                                    </div>
                                    <div class="drop-icon-card-ghost absolute inset-0 rounded-xl border-2 border-dashed border-indigo-400 opacity-0 z-0"></div>
                                </div>
                                <p class="text-sm font-semibold text-slate-700" id="label-template_gaji">Tarik file ke sini</p>
                                <p class="text-[11px] text-slate-500 mt-1">atau klik untuk menelusuri</p>
                                <div class="mt-3 px-2 py-0.5 rounded text-[10px] font-medium bg-slate-200/70 text-slate-500">.xls / .xlsx</div>
                            </div>
                            <input type="file" name="template_gaji" id="template_gaji" accept=".xls,.xlsx" class="hidden">
                        </div>
                        @error('template_gaji')
                            <p class="mt-2 text-xs font-medium text-red-500 flex items-center gap-1"><svg class="w-3 h-3" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 8v4m0 4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z" /></svg>{{ $message }}</p>
                        @enderror
                        <div id="chip-template_gaji" class="hidden mt-2 file-chip flex items-center gap-2 bg-indigo-50 border border-indigo-100 rounded-lg px-3 py-2">
                            <svg class="w-4 h-4 text-indigo-500 shrink-0" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24">
                                <path stroke-linecap="round" stroke-linejoin="round" d="M9 12.75L11.25 15 15 9.75M21 12a9 9 0 11-18 0 9 9 0 0118 0z"/>
                            </svg>
                            <span class="text-xs font-medium text-indigo-700 truncate" id="chipname-template_gaji"></span>
                        </div>
                    </div>

                </div>

                {{-- Submit Area Container --}}
                <div class="bg-slate-50 rounded-xl p-4 sm:p-5 flex flex-col sm:flex-row items-center justify-between gap-4 border border-slate-100">
                    
                    {{-- Progress indicator --}}
                    <div class="flex items-center gap-3 w-full sm:w-auto">
                        <div class="flex items-center gap-1.5 bg-white px-3 py-2 rounded-lg border border-slate-200 shadow-sm">
                            <div class="w-2 h-2 rounded-full transition-colors duration-300 bg-slate-200" id="dot-standard_report"></div>
                            <div class="w-2 h-2 rounded-full transition-colors duration-300 bg-slate-200" id="dot-keterangan_lain"></div>
                            <div class="w-2 h-2 rounded-full transition-colors duration-300 bg-slate-200" id="dot-template_gaji"></div>
                        </div>
                        <div>
                            <p class="text-[13px] font-semibold text-slate-700" id="progress-label">0 dari 3 file dipilih</p>
                            <p class="text-[11px] text-slate-500">File tidak disimpan di server</p>
                        </div>
                    </div>

                    {{-- Submit button --}}
                    <button
                        type="submit"
                        id="btn-proses"
                        class="w-full sm:w-auto inline-flex items-center justify-center gap-2.5 bg-slate-900 hover:bg-slate-800 active:bg-slate-950 text-white text-sm font-semibold px-6 py-3 rounded-lg shadow-md shadow-slate-900/10 transition-all duration-200 disabled:opacity-60 disabled:cursor-not-allowed whitespace-nowrap"
                    >
                        <svg id="btn-icon" class="w-4 h-4" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24">
                            <path stroke-linecap="round" stroke-linejoin="round" d="M3 16.5v2.25A2.25 2.25 0 005.25 21h13.5A2.25 2.25 0 0021 18.75V16.5M12 3v13.5m0 0l-4.5-4.5M12 16.5l4.5-4.5"/>
                        </svg>
                        <span id="btn-text">Proses &amp; Unduh</span>
                    </button>
                </div>
            </form>

            {{-- Info notice --}}
            <div class="mt-8 flex items-start gap-3 rounded-xl border border-amber-200/60 bg-amber-50/50 px-4 py-4">
                <div class="bg-amber-100 p-1.5 rounded-full shrink-0">
                    <svg class="w-4 h-4 text-amber-600" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24">
                        <path stroke-linecap="round" stroke-linejoin="round" d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-3L13.732 4c-.77-1.333-2.694-1.333-3.464 0L3.34 16c-.77 1.333.192 3 1.732 3z"/>
                    </svg>
                </div>
                <div>
                    <h4 class="text-sm font-bold text-amber-800">Catatan Penting</h4>
                    <p class="mt-1 text-xs text-amber-700/80 leading-relaxed">
                        Kolom <strong>No Rekening, Jabatan, Area, Gaji Pokok, Bonus</strong> tidak diisi otomatis — lengkapi secara manual setelah file diunduh. Kolom <em>Catatan</em> adalah draf dari Keterangan Lain, harap verifikasi sebelum diteruskan ke bagian payroll.
                    </p>
                </div>
            </div>

        </div>

        {{-- Footer text --}}
        <div class="mt-8 text-center">
            <p class="text-[11px] font-medium text-slate-400">&copy; {{ date('Y') }} All rights reserved.</p>
        </div>

    </main>

    {{-- Toast notification --}}
    <div id="toast-success" class="fixed bottom-6 sm:bottom-auto sm:top-6 right-1/2 translate-x-1/2 sm:translate-x-0 sm:right-6 z-50 hidden">
        <div class="flex items-center gap-3 bg-slate-900 text-white px-5 py-3.5 rounded-xl shadow-2xl shadow-emerald-900/20 border border-slate-800" style="animation: fadeSlideIn 0.4s cubic-bezier(0.16, 1, 0.3, 1) forwards;">
            <div class="bg-emerald-500/20 p-1.5 rounded-full">
                <svg class="w-4 h-4 text-emerald-400 shrink-0" fill="none" stroke="currentColor" stroke-width="3" viewBox="0 0 24 24">
                    <path stroke-linecap="round" stroke-linejoin="round" d="M4.5 12.75l6 6 9-13.5" />
                </svg>
            </div>
            <span class="text-sm font-semibold tracking-wide">Rekap berhasil diunduh!</span>
        </div>
    </div>

    <div id="toast-error" class="fixed bottom-6 sm:bottom-auto sm:top-6 right-1/2 translate-x-1/2 sm:translate-x-0 sm:right-6 z-50 hidden w-[90%] sm:w-auto max-w-md">
        <div class="flex items-start gap-3 bg-red-600 text-white px-5 py-4 rounded-xl shadow-2xl shadow-red-900/20 border border-red-500" style="animation: fadeSlideIn 0.4s cubic-bezier(0.16, 1, 0.3, 1) forwards;">
            <div class="bg-white/20 p-1.5 rounded-full shrink-0 mt-0.5">
                <svg class="w-4 h-4 text-white shrink-0" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24">
                    <path stroke-linecap="round" stroke-linejoin="round" d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-3L13.732 4c-.77-1.333-2.694-1.333-3.464 0L3.34 16c-.77 1.333.192 3 1.732 3z"/>
                </svg>
            </div>
            <span class="text-sm font-medium leading-relaxed" id="toast-error-text">Gagal memproses file.</span>
        </div>
    </div>

    <script>
    (function () {
        var fields = ['standard_report', 'keterangan_lain', 'template_gaji'];

        // ── Drag & Drop setup ──────────────────────────────────────────────
        fields.forEach(function (name) {
            var zone  = document.getElementById('zone-' + name);
            var input = document.getElementById(name);

            // Klik pada zona → buka dialog file
            zone.addEventListener('click', function () { input.click(); });

            // Prevent default bawaan browser untuk drag
            ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(function (ev) {
                zone.addEventListener(ev, function (e) {
                    e.preventDefault();
                    e.stopPropagation();
                });
            });

            // Highlight saat drag masuk
            zone.addEventListener('dragenter', function () { zone.classList.add('drag-over'); });
            zone.addEventListener('dragover',  function () { zone.classList.add('drag-over'); });
            zone.addEventListener('dragleave', function () { zone.classList.remove('drag-over'); });

            // Drop file — kunci: assign ke input.files via DataTransfer agar ikut terbawa form submit
            zone.addEventListener('drop', function (e) {
                zone.classList.remove('drag-over');
                var file = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0];
                if (!file) { return; }

                // Assign file ke elemen <input> agar browser menyertakannya saat submit
                var dt = new DataTransfer();
                dt.items.add(file);
                input.files = dt.files;

                assignFile(name, file);
            });

            // Pilih via dialog file
            input.addEventListener('change', function () {
                if (input.files && input.files[0]) {
                    assignFile(name, input.files[0]);
                }
            });
        });

        function assignFile(name, file) {
            // Tampilkan nama file di label zona
            var label = document.getElementById('label-' + name);
            if (label) {
                label.textContent = file.name.length > 22
                    ? file.name.substring(0, 20) + '…'
                    : file.name;
            }

            // Tampilkan chip di bawah zona
            var chip     = document.getElementById('chip-' + name);
            var chipName = document.getElementById('chipname-' + name);
            if (chip && chipName) {
                chipName.textContent = file.name;
                chip.classList.remove('hidden');
            }

            // Tandai zona sudah ada file
            var zone = document.getElementById('zone-' + name);
            zone.classList.add('has-file');
            zone.classList.remove('border-red-400', 'bg-red-50');

            updateProgress();
        }

        // ── Progress dots ──────────────────────────────────────────────────
        function updateProgress() {
            var count = 0;
            fields.forEach(function (name) {
                var dot   = document.getElementById('dot-' + name);
                var input = document.getElementById(name);
                var hasFile = input.files && input.files.length > 0;

                if (hasFile) {
                    count++;
                    dot.classList.remove('bg-slate-200');
                    dot.classList.add('bg-indigo-500');
                } else {
                    dot.classList.remove('bg-indigo-500');
                    dot.classList.add('bg-slate-200');
                }
            });
            document.getElementById('progress-label').textContent =
                count + ' dari 3 file dipilih';
        }

        // ── Reset form ke kondisi awal ─────────────────────────────────────
        function resetForm() {
            // Reset semua file input
            fields.forEach(function (name) {
                var input = document.getElementById(name);
                var zone  = document.getElementById('zone-' + name);
                var label = document.getElementById('label-' + name);
                var chip  = document.getElementById('chip-' + name);

                // Kosongkan input file
                input.value = '';

                // Reset zona
                zone.classList.remove('has-file', 'border-red-400', 'bg-red-50');
                zone.classList.add('border-slate-200');

                // Reset label
                if (label) {
                    label.textContent = 'Tarik file ke sini';
                }

                // Sembunyikan chip
                if (chip) {
                    chip.classList.add('hidden');
                }
            });

            updateProgress();
        }

        // ── Tombol kembali ke state normal ──────────────────────────────────
        function resetButton() {
            var btn  = document.getElementById('btn-proses');
            var icon = document.getElementById('btn-icon');
            var text = document.getElementById('btn-text');
            btn.disabled = false;
            icon.classList.remove('hidden');
            text.innerHTML = 'Proses &amp; Unduh';
        }

        // ── Toast notifications ────────────────────────────────────────────
        function showToast(type, message) {
            var toastId = type === 'success' ? 'toast-success' : 'toast-error';
            var toast = document.getElementById(toastId);

            if (type === 'error' && message) {
                document.getElementById('toast-error-text').textContent = message;
            }

            toast.classList.remove('hidden');
            setTimeout(function () {
                toast.classList.add('hidden');
            }, 4500);
        }

        // ── Form submit via AJAX (fetch) ───────────────────────────────────
        document.getElementById('form-absensi').addEventListener('submit', function (e) {
            e.preventDefault(); // Selalu prevent — kita handle semuanya via fetch

            var firstInvalid = null;

            fields.forEach(function (name) {
                var input = document.getElementById(name);
                var zone  = document.getElementById('zone-' + name);
                var hasFile = input.files && input.files.length > 0;

                if (!hasFile) {
                    zone.classList.add('border-red-400', 'bg-red-50');
                    zone.classList.remove('border-slate-200');
                    if (!firstInvalid) { firstInvalid = zone; }
                } else {
                    zone.classList.remove('border-red-400', 'bg-red-50');
                }
            });

            if (firstInvalid) {
                firstInvalid.scrollIntoView({ behavior: 'smooth', block: 'center' });
                return;
            }

            // Aktifkan loading state
            var btn  = document.getElementById('btn-proses');
            var icon = document.getElementById('btn-icon');
            var text = document.getElementById('btn-text');
            btn.disabled = true;
            icon.classList.add('hidden');
            text.innerHTML = 'Memproses...<span class="spinner"></span>';

            // Bangun FormData dari form
            var form = document.getElementById('form-absensi');
            var formData = new FormData(form);

            fetch(form.action, {
                method: 'POST',
                body: formData,
                headers: {
                    'X-Requested-With': 'XMLHttpRequest',
                    'Accept': 'application/octet-stream, application/json',
                },
            })
            .then(function (response) {
                if (!response.ok) {
                    // Coba parse sebagai JSON untuk pesan error
                    return response.json().then(function (data) {
                        // Laravel validation error format: {message: "...", errors: {...}}
                        if (data.errors) {
                            var msgs = [];
                            for (var field in data.errors) {
                                msgs = msgs.concat(data.errors[field]);
                            }
                            throw new Error(msgs.join('\n'));
                        }
                        throw new Error(data.error || data.message || 'Gagal memproses file.');
                    }).catch(function (jsonErr) {
                        // Jika error sudah di-throw dari blok atas, teruskan
                        if (jsonErr instanceof Error) {
                            throw jsonErr;
                        }
                        throw new Error('Gagal memproses file. Pastikan format dan isi file sesuai, lalu coba kembali.');
                    });
                }

                // Ambil nama file dari header response
                var filename = 'Rekap Absensi.xlsx';
                var disposition = response.headers.get('Content-Disposition');
                if (disposition) {
                    var filenameMatch = disposition.match(/filename[^;=\n]*=["']?([^"';\n]*)["']?/);
                    if (filenameMatch && filenameMatch[1]) {
                        filename = decodeURIComponent(filenameMatch[1].replace(/['"]/g, ''));
                    }
                }
                var customFilename = response.headers.get('X-Download-Filename');
                if (customFilename) {
                    filename = customFilename;
                }

                return response.blob().then(function (blob) {
                    return { blob: blob, filename: filename };
                });
            })
            .then(function (result) {
                // Trigger download manual via blob URL
                var url = window.URL.createObjectURL(result.blob);
                var a = document.createElement('a');
                a.href = url;
                a.download = result.filename;
                document.body.appendChild(a);
                a.click();
                document.body.removeChild(a);
                window.URL.revokeObjectURL(url);

                // Sukses: reset button, reset form, tampilkan toast
                resetButton();
                resetForm();
                showToast('success');
            })
            .catch(function (error) {
                resetButton();
                showToast('error', error.message);
            });
        });

    })();
    </script>
</body>
</html>