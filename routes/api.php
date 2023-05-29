<?php

use Illuminate\Http\Request;
use Illuminate\Support\Facades\Route;
use App\Http\Controllers\PrintWordDocumentController;
use App\Http\Controllers\PrintPdfDocumentController;
use App\Http\Controllers\ReadPdfDocumentController;

/*
|--------------------------------------------------------------------------
| API Routes
|--------------------------------------------------------------------------
|
| Here is where you can register API routes for your application. These
| routes are loaded by the RouteServiceProvider within a group which
| is assigned the "api" middleware group. Enjoy building your API!
|
*/

Route::middleware('auth:sanctum')->get('/user', function (Request $request) {
    return $request->user();
});

Route::post('/print-word', [PrintWordDocumentController::class, 'print']);
Route::post('/print-pdf', [PrintPdfDocumentController::class, 'print']);

