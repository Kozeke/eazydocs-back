<?php

namespace App\Http\Controllers;

use App\DTO\PrintPDFRecordDTO;
use App\Http\Requests\PdfPrintRequest;
use App\Http\Services\PrintPDFDocument\PrintRecordService;
use App\Http\Services\PrintPDFDocument\PrintTableService;
use Illuminate\Support\Collection as SupportCollection;
use Mpdf\Mpdf;
use Mpdf\MpdfException;
use Spatie\DataTransferObject\Exceptions\UnknownProperties;

class PrintPDFDocumentController extends Controller
{
    /**
     * Class ActEditorPrintWordService
     * @package App\Contollers
     *
     * @author Kozy-Korpesh Tolep
     */


    /**
     * @var int
     */
    private const DEFAULT_INDENT = 10;

    /**
     * @var int
     */
    private const PAGE_WIDTH_IN_MM = 210;


    /**
     * @var array
     */
    private array $predefinedStyles = [
        //Стили разделов:
        'styleSection' => [
            'marginLeft' => 2.5,
            'marginRight' => 2.5,
            'marginTop' => 2.5,
            'marginBottom' => 2.5,
        ],
        'fontSize' => 12,
        'fontFamily' => 'Times New Roman',
        'substringFontSize' => 8
    ];

    /**
     * @var  Mpdf
     */
    private Mpdf $pdf;

    /**
     * @var int|float
     */
    private int|float $pageWidthInMmWithoutMargins;

    /**
     * @param PdfPrintRequest $request
     * @return string|null
     * @throws MpdfException
     * @throws UnknownProperties
     */
    public function print(PdfPrintRequest $request): ?string
    {
        $docStyleSettings = $request->input('documentDefaultSettings');
        $documentLines = $request->input('documentLines');
        $leftIndent = $docStyleSettings->leftIndent ? $docStyleSettings->leftIndent * self::DEFAULT_INDENT : self::DEFAULT_INDENT;
        $rightIndent = $docStyleSettings->rightIndent ? $docStyleSettings->rightIndent * self::DEFAULT_INDENT : self::DEFAULT_INDENT;
        $topIndent = $docStyleSettings->topIndent ? $docStyleSettings->topIndent * self::DEFAULT_INDENT : self::DEFAULT_INDENT;
        $bottomIndent = $docStyleSettings->bottomIndent ? $docStyleSettings->bottomIndent * self::DEFAULT_INDENT : self::DEFAULT_INDENT;
        $this->pageWidthInMmWithoutMargins = self::PAGE_WIDTH_IN_MM - $leftIndent - $rightIndent;
        $this->pdf = new Mpdf([
            'tempDir' => storage_path('app/tempDirMpdf'),
            'margin_left' => $leftIndent,
            'margin_right' => $rightIndent,
            'margin_top' => $topIndent,
            'margin_bottom' => $bottomIndent,
        ]);
        $this->pdf->SetAutoPageBreak(true, $bottomIndent);

        $this->printDocumentLines($documentLines);
        return $this->pdf->Output(str_replace("\n", "", $docStyleSettings->name) . '.pdf', 'D');
    }

    /**
     * @throws MpdfException
     * @throws UnknownProperties
     */
    private function printDocumentLines(SupportCollection $documentLines)
    {
        foreach ($documentLines as $key => $documentLine) {
            if ($documentLine['viewTitleOnPrint'] || $documentLine['viewFieldOnPrint']) {
                if ($documentLine['tab_type'] === 'record') {
                    if ($documentLine['line_type'] === 'record') {
                        (new PrintRecordService())->ifTabTypeRecord(
                            PrintPDFRecordDTO::createFromArray(
                                [
                                    "key_of_current_document_line" => $key,
                                    "pdf" => $this->pdf,
                                    "page_width_in_mm_without_margins" => $this->pageWidthInMmWithoutMargins,
                                    "document_lines" => $documentLine,
                                ]
                            )
                        );
                    }
                }
                if ($documentLine['tab_type'] === 'table') {
                    $this->pdf->Ln($this->predefinedStyles['spacing']);
                    (new PrintTableService())->printAsTable(
                        PrintPDFRecordDTO::createFromArray(
                            [
                                "pdf" => $this->pdf,
                                "page_width_in_mm_without_margins" => $this->pageWidthInMmWithoutMargins,
                                "document_lines" => $documentLine,
                            ]
                        )
                    );
                }
            }
        }
    }
}
