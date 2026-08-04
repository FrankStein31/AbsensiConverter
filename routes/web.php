<?php

use App\Http\Controllers\AbsensiController;
use Illuminate\Support\Facades\Route;

Route::get('/', [AbsensiController::class, 'index'])->name('absensi.index');
Route::post('/', [AbsensiController::class, 'proses'])->name('absensi.proses');
