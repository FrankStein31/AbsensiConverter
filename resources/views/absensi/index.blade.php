<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta name="robots" content="noindex, nofollow">
    <title>Konverter Rekap Absensi — Scomptec</title>
    @if (file_exists(public_path('build/manifest.json')) || file_exists(public_path('hot')))
        @vite(['resources/css/app.css', 'resources/js/app.js'])
    @else
        <script src="https://unpkg.com/@tailwindcss/browser@4"></script>
    @endif
    <style>
        @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;500;600&display=swap');
        body { font-family: 'IBM Plex Sans', ui-sans-serif, system-ui, sans-serif; }

        /* Grid dot pattern background */
        .grid-pattern {
            background-image: radial-gradient(circle, #cbd5e1 1px, transparent 1px);
            background-size: 20px 20px;
        }

        /* Drop zone states */
        .drop-zone { transition: all 0.2s ease; }
        .drop-zone.drag-over {
            border-color: #334155 !important;
            background-color: #f1f5f9 !important;
        }
        .drop-zone.drag-over .drop-icon-card {
            transform: translate(8px, -8px);
            box-shadow: 0 20px 40px rgba(0,0,0,0.15);
        }
        .drop-zone.drag-over .drop-icon-card-ghost {
            opacity: 1 !important;
        }
        .drop-zone.has-file {
            border-color: #94a3b8;
            background-color: #f8fafc;
        }

        /* Icon card animation */
        .drop-icon-card { transition: transform 0.2s ease, box-shadow 0.2s ease; }
        .drop-icon-card-ghost { transition: opacity 0.2s ease; }

        /* Loading spinner */
        @keyframes spin { to { transform: rotate(360deg); } }
        .spinner {
            display: inline-block;
            width: 14px; height: 14px;
            border: 2px solid rgba(255,255,255,0.35);
            border-top-color: white;
            border-radius: 50%;
            animation: spin 0.6s linear infinite;
            vertical-align: middle;
            margin-left: 6px;
        }

        /* File chip fade-in */
        @keyframes fadeSlideIn {
            from { opacity: 0; transform: translateY(6px); }
            to   { opacity: 1; transform: translateY(0); }
        }
        .file-chip { animation: fadeSlideIn 0.2s ease forwards; }
    </style>
</head>
<body class="bg-slate-100 min-h-screen antialiased">

    {{-- Header --}}
    <header class="bg-white border-b border-slate-200 px-5 py-2.5 flex items-center justify-between">
        <div class="flex items-center gap-2.5">
            <div class="w-5 h-5 bg-slate-800 rounded-sm flex items-center justify-center shrink-0">
                <svg class="w-3 h-3 text-white" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24">
                    <path stroke-linecap="round" stroke-linejoin="round" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"/>
                </svg>
            </div>
            <span class="text-sm font-semibold text-slate-800 tracking-tight">Konverter Rekap Absensi</span>
        </div>
        <span class="text-xs text-slate-400 hidden sm:block">Data tidak disimpan di server &bull; Scomptec</span>
    </header>

    <main class="max-w-2xl mx-auto px-4 py-8">

        {{-- Page Heading --}}
        <div class="mb-6">
            <h1 class="text-base font-semibold text-slate-800">Upload File Absensi</h1>
            <p class="mt-1 text-xs text-slate-500 leading-relaxed max-w-lg">
                Upload 3 file Excel: Standard Report mesin fingerprint, Keterangan Lain (sakit/cuti/izin), dan Template Gaji.
                Sistem akan menghasilkan satu file rekap yang siap diunduh.
            </p>
        </div>

        {{-- Error / Validation Alerts --}}
        @if (session('error'))
            <div class="mb-5 flex items-start gap-2.5 rounded-sm border border-red-200 bg-red-50 px-3 py-2.5">
                <svg class="w-4 h-4 text-red-500 mt-0.5 shrink-0" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24">
                    <path stroke-linecap="round" stroke-linejoin="round" d="M12 9v3.75m-9.303 3.376c-.866 1.5.217 3.374 1.948 3.374h14.71c1.73 0 2.813-1.874 1.948-3.374L13.949 3.378c-.866-1.5-3.032-1.5-3.898 0L2.697 16.126zM12 15.75h.007v.008H12v-.008z"/>
                </svg>
                <div>
                    <p class="text-xs font-semibold text-red-700">Gagal Memproses</p>
                    <p class="text-xs text-red-600 mt-0.5">{{ session('error') }}</p>
                </div>
            </div>
        @endif

        @if ($errors->any())
            <div class="mb-5 rounded-sm border border-red-200 bg-red-50 px-3 py-2.5">
                <p class="text-xs font-semibold text-red-700 mb-1.5">Periksa input berikut:</p>
                <ul class="space-y-0.5">
                    @foreach ($errors->all() as $error)
                        <li class="text-xs text-red-600 flex items-center gap-1.5">
                            <span class="w-1 h-1 bg-red-400 rounded-full shrink-0"></span>{{ $error }}
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

            <div class="grid grid-cols-1 sm:grid-cols-3 gap-4 mb-5">

                {{-- Drop Zone 1: Standard Report --}}
                <div>
                    <p class="text-xs font-medium text-slate-600 mb-1.5 flex items-center gap-1.5">
                        <span class="inline-flex items-center justify-center w-4 h-4 bg-slate-800 text-white rounded-full text-[9px] font-bold shrink-0">1</span>
                        Standard Report
                    </p>
                    <div
                        class="drop-zone relative rounded-md border-2 border-dashed border-slate-300 bg-white cursor-pointer overflow-hidden"
                        data-target="standard_report"
                        id="zone-standard_report"
                    >
                        {{-- Grid dot bg --}}
                        <div class="grid-pattern absolute inset-0 opacity-40 pointer-events-none"></div>
                        <div class="relative z-10 flex flex-col items-center justify-center py-8 px-3 text-center">
                            {{-- Animated icon card --}}
                            <div class="relative w-16 h-14 mb-3">
                                <div class="drop-icon-card absolute inset-0 bg-white rounded-md shadow-md flex items-center justify-center z-10">
                                    <svg class="w-5 h-5 text-slate-400" fill="none" stroke="currentColor" stroke-width="1.5" viewBox="0 0 24 24">
                                        <path stroke-linecap="round" stroke-linejoin="round" d="M19.5 14.25v-2.625a3.375 3.375 0 00-3.375-3.375h-1.5A1.125 1.125 0 0113.5 7.125v-1.5a3.375 3.375 0 00-3.375-3.375H8.25m2.25 0H5.625c-.621 0-1.125.504-1.125 1.125v17.25c0 .621.504 1.125 1.125 1.125h12.75c.621 0 1.125-.504 1.125-1.125V11.25a9 9 0 00-9-9z"/>
                                    </svg>
                                </div>
                                <div class="drop-icon-card-ghost absolute inset-0 rounded-md border-2 border-dashed border-sky-400 opacity-0 z-0"></div>
                            </div>
                            <p class="text-xs font-medium text-slate-600" id="label-standard_report">Drag file di sini</p>
                            <p class="text-[10px] text-slate-400 mt-0.5">atau klik untuk pilih</p>
                            <p class="text-[10px] text-slate-300 mt-1">.xls / .xlsx</p>
                        </div>
                        <input type="file" name="standard_report" id="standard_report" accept=".xls,.xlsx" class="hidden">
                    </div>
                    @error('standard_report')
                        <p class="mt-1 text-xs text-red-500">{{ $message }}</p>
                    @enderror
                    {{-- File chip --}}
                    <div id="chip-standard_report" class="hidden mt-1.5 file-chip flex items-center gap-1.5 bg-slate-100 border border-slate-200 rounded px-2 py-1">
                        <svg class="w-3 h-3 text-slate-400 shrink-0" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24">
                            <path stroke-linecap="round" stroke-linejoin="round" d="M9 12.75L11.25 15 15 9.75M21 12a9 9 0 11-18 0 9 9 0 0118 0z"/>
                        </svg>
                        <span class="text-xs text-slate-600 truncate" id="chipname-standard_report"></span>
                    </div>
                </div>

                {{-- Drop Zone 2: Keterangan Lain --}}
                <div>
                    <p class="text-xs font-medium text-slate-600 mb-1.5 flex items-center gap-1.5">
                        <span class="inline-flex items-center justify-center w-4 h-4 bg-slate-800 text-white rounded-full text-[9px] font-bold shrink-0">2</span>
                        Keterangan Lain
                    </p>
                    <div
                        class="drop-zone relative rounded-md border-2 border-dashed border-slate-300 bg-white cursor-pointer overflow-hidden"
                        data-target="keterangan_lain"
                        id="zone-keterangan_lain"
                    >
                        <div class="grid-pattern absolute inset-0 opacity-40 pointer-events-none"></div>
                        <div class="relative z-10 flex flex-col items-center justify-center py-8 px-3 text-center">
                            <div class="relative w-16 h-14 mb-3">
                                <div class="drop-icon-card absolute inset-0 bg-white rounded-md shadow-md flex items-center justify-center z-10">
                                    <svg class="w-5 h-5 text-slate-400" fill="none" stroke="currentColor" stroke-width="1.5" viewBox="0 0 24 24">
                                        <path stroke-linecap="round" stroke-linejoin="round" d="M19.5 14.25v-2.625a3.375 3.375 0 00-3.375-3.375h-1.5A1.125 1.125 0 0113.5 7.125v-1.5a3.375 3.375 0 00-3.375-3.375H8.25m2.25 0H5.625c-.621 0-1.125.504-1.125 1.125v17.25c0 .621.504 1.125 1.125 1.125h12.75c.621 0 1.125-.504 1.125-1.125V11.25a9 9 0 00-9-9z"/>
                                    </svg>
                                </div>
                                <div class="drop-icon-card-ghost absolute inset-0 rounded-md border-2 border-dashed border-sky-400 opacity-0 z-0"></div>
                            </div>
                            <p class="text-xs font-medium text-slate-600" id="label-keterangan_lain">Drag file di sini</p>
                            <p class="text-[10px] text-slate-400 mt-0.5">atau klik untuk pilih</p>
                            <p class="text-[10px] text-slate-300 mt-1">.xls / .xlsx</p>
                        </div>
                        <input type="file" name="keterangan_lain" id="keterangan_lain" accept=".xls,.xlsx" class="hidden">
                    </div>
                    @error('keterangan_lain')
                        <p class="mt-1 text-xs text-red-500">{{ $message }}</p>
                    @enderror
                    <div id="chip-keterangan_lain" class="hidden mt-1.5 file-chip flex items-center gap-1.5 bg-slate-100 border border-slate-200 rounded px-2 py-1">
                        <svg class="w-3 h-3 text-slate-400 shrink-0" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24">
                            <path stroke-linecap="round" stroke-linejoin="round" d="M9 12.75L11.25 15 15 9.75M21 12a9 9 0 11-18 0 9 9 0 0118 0z"/>
                        </svg>
                        <span class="text-xs text-slate-600 truncate" id="chipname-keterangan_lain"></span>
                    </div>
                </div>

                {{-- Drop Zone 3: Template Gaji --}}
                <div>
                    <p class="text-xs font-medium text-slate-600 mb-1.5 flex items-center gap-1.5">
                        <span class="inline-flex items-center justify-center w-4 h-4 bg-slate-800 text-white rounded-full text-[9px] font-bold shrink-0">3</span>
                        Template Gaji
                    </p>
                    <div
                        class="drop-zone relative rounded-md border-2 border-dashed border-slate-300 bg-white cursor-pointer overflow-hidden"
                        data-target="template_gaji"
                        id="zone-template_gaji"
                    >
                        <div class="grid-pattern absolute inset-0 opacity-40 pointer-events-none"></div>
                        <div class="relative z-10 flex flex-col items-center justify-center py-8 px-3 text-center">
                            <div class="relative w-16 h-14 mb-3">
                                <div class="drop-icon-card absolute inset-0 bg-white rounded-md shadow-md flex items-center justify-center z-10">
                                    <svg class="w-5 h-5 text-slate-400" fill="none" stroke="currentColor" stroke-width="1.5" viewBox="0 0 24 24">
                                        <path stroke-linecap="round" stroke-linejoin="round" d="M19.5 14.25v-2.625a3.375 3.375 0 00-3.375-3.375h-1.5A1.125 1.125 0 0113.5 7.125v-1.5a3.375 3.375 0 00-3.375-3.375H8.25m2.25 0H5.625c-.621 0-1.125.504-1.125 1.125v17.25c0 .621.504 1.125 1.125 1.125h12.75c.621 0 1.125-.504 1.125-1.125V11.25a9 9 0 00-9-9z"/>
                                    </svg>
                                </div>
                                <div class="drop-icon-card-ghost absolute inset-0 rounded-md border-2 border-dashed border-sky-400 opacity-0 z-0"></div>
                            </div>
                            <p class="text-xs font-medium text-slate-600" id="label-template_gaji">Drag file di sini</p>
                            <p class="text-[10px] text-slate-400 mt-0.5">atau klik untuk pilih</p>
                            <p class="text-[10px] text-slate-300 mt-1">.xls / .xlsx</p>
                        </div>
                        <input type="file" name="template_gaji" id="template_gaji" accept=".xls,.xlsx" class="hidden">
                    </div>
                    @error('template_gaji')
                        <p class="mt-1 text-xs text-red-500">{{ $message }}</p>
                    @enderror
                    <div id="chip-template_gaji" class="hidden mt-1.5 file-chip flex items-center gap-1.5 bg-slate-100 border border-slate-200 rounded px-2 py-1">
                        <svg class="w-3 h-3 text-slate-400 shrink-0" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24">
                            <path stroke-linecap="round" stroke-linejoin="round" d="M9 12.75L11.25 15 15 9.75M21 12a9 9 0 11-18 0 9 9 0 0118 0z"/>
                        </svg>
                        <span class="text-xs text-slate-600 truncate" id="chipname-template_gaji"></span>
                    </div>
                </div>

            </div>

            {{-- Progress indicator --}}
            <div class="mb-4 flex items-center gap-2">
                <div class="flex items-center gap-1">
                    <div class="w-2 h-2 rounded-full bg-slate-200" id="dot-standard_report"></div>
                    <div class="w-2 h-2 rounded-full bg-slate-200" id="dot-keterangan_lain"></div>
                    <div class="w-2 h-2 rounded-full bg-slate-200" id="dot-template_gaji"></div>
                </div>
                <p class="text-xs text-slate-400" id="progress-label">0 dari 3 file dipilih</p>
            </div>

            {{-- Submit button --}}
            <div class="flex items-center justify-between gap-4 bg-white border border-slate-200 rounded-sm px-4 py-3">
                <p class="text-xs text-slate-400">File tidak disimpan di server setelah diproses.</p>
                <button
                    type="submit"
                    id="btn-proses"
                    class="inline-flex items-center gap-2 bg-slate-800 hover:bg-slate-700 active:bg-slate-900 text-white text-xs font-medium px-4 py-2 rounded-sm transition-colors duration-150 disabled:opacity-50 disabled:cursor-not-allowed whitespace-nowrap"
                >
                    <svg id="btn-icon" class="w-3.5 h-3.5" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24">
                        <path stroke-linecap="round" stroke-linejoin="round" d="M3 16.5v2.25A2.25 2.25 0 005.25 21h13.5A2.25 2.25 0 0021 18.75V16.5M12 3v13.5m0 0l-4.5-4.5M12 16.5l4.5-4.5"/>
                    </svg>
                    <span id="btn-text">Proses &amp; Unduh</span>
                </button>
            </div>
        </form>

        {{-- Info notice --}}
        <div class="mt-4 flex items-start gap-2 rounded-sm border border-amber-200 bg-amber-50 px-3 py-2.5">
            <svg class="w-3.5 h-3.5 text-amber-500 mt-0.5 shrink-0" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24">
                <path stroke-linecap="round" stroke-linejoin="round" d="M11.25 11.25l.041-.02a.75.75 0 011.063.852l-.708 2.836a.75.75 0 001.063.853l.041-.021M21 12a9 9 0 11-18 0 9 9 0 0118 0zm-9-3.75h.008v.008H12V8.25z"/>
            </svg>
            <p class="text-xs text-slate-600 leading-relaxed">
                Kolom <strong>No Rekening, Jabatan, Area, Gaji Pokok, Bonus</strong> tidak diisi otomatis — lengkapi manual setelah unduh.
                Kolom <em>Catatan</em> adalah draf dari Keterangan Lain, verifikasi sebelum dikirim ke payroll.
            </p>
        </div>

    </main>

    <script>
    (function () {
        const fields = ['standard_report', 'keterangan_lain', 'template_gaji'];
        const fileMap = {};  // nama field -> File object

        // ── Drag & Drop setup ──────────────────────────────────────────────
        fields.forEach(function (name) {
            const zone  = document.getElementById('zone-' + name);
            const input = document.getElementById(name);

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
                    dot.classList.add('bg-slate-700');
                } else {
                    dot.classList.remove('bg-slate-700');
                    dot.classList.add('bg-slate-200');
                }
            });
            document.getElementById('progress-label').textContent =
                count + ' dari 3 file dipilih';
        }

        // ── Form submit ────────────────────────────────────────────────────
        document.getElementById('form-absensi').addEventListener('submit', function (e) {
            var firstInvalid = null;

            fields.forEach(function (name) {
                var input = document.getElementById(name);
                var zone  = document.getElementById('zone-' + name);
                var hasFile = input.files && input.files.length > 0;

                if (!hasFile) {
                    zone.classList.add('border-red-400', 'bg-red-50');
                    zone.classList.remove('border-slate-300');
                    if (!firstInvalid) { firstInvalid = zone; }
                } else {
                    zone.classList.remove('border-red-400', 'bg-red-50');
                }
            });

            if (firstInvalid) {
                e.preventDefault();
                firstInvalid.scrollIntoView({ behavior: 'smooth', block: 'center' });
                return;
            }

            // Double-submit prevention
            var btn  = document.getElementById('btn-proses');
            var icon = document.getElementById('btn-icon');
            var text = document.getElementById('btn-text');
            btn.disabled = true;
            icon.classList.add('hidden');
            text.innerHTML = 'Memproses...<span class="spinner"></span>';
        });

    })();
    </script>
</body>
</html>