<?php

namespace App\Http\Controllers;


use App\DTO\PrintWordRecordDTO;
use App\DTO\PrintWordTableDTO;
use App\Extensions\PrintWordDocument\PrintWordTrait;
use App\Http\Services\PrintWordDocument\PrintRecordService;
use App\Http\Services\PrintWordDocument\PrintTableService;
use Illuminate\Support\Facades\Storage;
use PhpOffice\PhpWord\Element\Section;
use PhpOffice\PhpWord\Element\Table;
use PhpOffice\PhpWord\Exception\Exception;
use PhpOffice\PhpWord\IOFactory;
use PhpOffice\PhpWord\PhpWord;
use PhpOffice\PhpWord\Shared\Converter;
use Spatie\DataTransferObject\Exceptions\UnknownProperties;
use Symfony\Component\HttpFoundation\BinaryFileResponse;
use App\Http\Requests\PdfPrintRequest;

/**
 * Class ActEditorPrintWordService
 * @package App\Contollers
 *
 * @author Kozy-Korpesh Tolep
 */
class PrintWordDocumentController extends Controller
{

    use PrintWordTrait;

    /**
     * @var string
     */
    private const TABLE_STYLE_NAME_FOR_ACTS = 'Раздел 1';


    /**
     * @var float
     */
    private float $pageWidthInTwipWithoutMargins;

    /**
     * @var Table
     */
    private Table $table;

    /**
     * @var Section
     */
    private Section $section;


    /**
     * @param PdfPrintRequest $request
     * @return BinaryFileResponse
     * @throws Exception
     * @throws UnknownProperties
     */
    public function print(PdfPrintRequest $request): BinaryFileResponse
    {
        $phpWord = new PhpWord();
        $phpWord->addParagraphStyle('StyleSubs', array('align' => 'center', 'spaceAfter' => 10));
        $documentLines = $request->input('documentLines');
        $docStyleSettings = $request->input('documentDefaultSettings');
        $this->setPredefinedStyles($phpWord, $docStyleSettings);
        $this->calculatePageWidthInTwipWithoutMargins();
        $phpWord = $this->printDocumentLineRecords($documentLines, $phpWord);
        $objWriter = IOFactory::createWriter($phpWord);
        $PathToTempFile = Storage::path(
            'public/document.docx'
        );
        $objWriter->save($PathToTempFile);
        return response()->download($PathToTempFile);
    }

    /**
     * @return void
     */
    private function calculatePageWidthInTwipWithoutMargins(): void
    {
        $pageWidthInTwip = $this->section->getStyle()->getPageSizeW();
        $marginLeftInTwip = $this->section->getStyle()->getMarginLeft();
        $marginRightInTwip = $this->section->getStyle()->getMarginRight();
        $this->pageWidthInTwipWithoutMargins = $pageWidthInTwip - $marginLeftInTwip - $marginRightInTwip;
    }

    /**
     * @param PhpWord $phpWord
     * @param array $docStyleSettings
     * @return void
     */
    private function setPredefinedStyles(PhpWord $phpWord, array $docStyleSettings): void
    {
        $this->predefinedStyles['styleSection']['marginLeft'] = Converter::cmToTwip(
            $docStyleSettings['leftIndentInCm'] ?? $this->predefinedStyles['styleSection']['marginLeft']
        );
        $this->predefinedStyles['styleSection']['marginRight'] = Converter::cmToTwip(
            $docStyleSettings['rightIndentInCm'] ?? $this->predefinedStyles['styleSection']['marginRight']
        );
        $this->predefinedStyles['styleSection']['marginTop'] = Converter::cmToTwip(
            $docStyleSettings['topIndentInCm'] ?? $this->predefinedStyles['styleSection']['marginTop']
        );
        $this->predefinedStyles['styleSection']['marginBottom'] = Converter::cmToTwip(
            $docStyleSettings['bottomIndentInCm'] ?? $this->predefinedStyles['styleSection']['marginBottom']
        );
        $this->predefinedStyles['spacing'] = $docStyleSettings['spacing'] ?? $this->predefinedStyles['spacing'];
        $this->predefinedStyles['fontSize'] = $docStyleSettings['fontSize'] ?? $this->predefinedStyles['fontSize'];
        $this->predefinedStyles['fontFamily'] = $docStyleSettings['fontFamily'] ?? $this->predefinedStyles['fontFamily'];
        $this->predefinedStyles['styleComments']['size'] = $docStyleSettings['substringFontSize'] ?? $this->predefinedStyles['styleComments']['size'];
        $phpWord->setDefaultFontName($this->predefinedStyles['fontFamily']);
        $phpWord->setDefaultFontSize($this->predefinedStyles['fontSize']);
    }


    /**
     * @param array $documentLines
     * @param PhpWord $phpWord
     * @return PhpWord
     * @throws UnknownProperties
     */
    private function printDocumentLineRecords(array $documentLines, PhpWord $phpWord): PhpWord
    {
        $this->titleHorizontalAlignmentOnPrint = '';
        $this->section = $phpWord->addSection($this->predefinedStyles["styleSection"]);
        $pageWidthInTwip = $this->section->getStyle()->getPageSizeW();
        $marginLeftInTwip = $this->section->getStyle()->getMarginLeft();
        $marginRightInTwip = $this->section->getStyle()->getMarginRight();
        $this->pageWidthInTwipWithoutMargins = $pageWidthInTwip - $marginLeftInTwip - $marginRightInTwip;
        foreach ($documentLines as $key => $documentLine) {
            if ($documentLine['line_type'] === 'table') {
                $phpWord->addTableStyle(self::TABLE_STYLE_NAME_FOR_ACTS, $this->predefinedStyles["styleTable"], []);
                (new PrintTableService())->printTable(
                    PrintWordTableDTO::createFromArray(
                        [
                            "section" => $this->section,
                            "page_width_in_twip_without_margins" => $this->pageWidthInTwipWithoutMargins,
                            "current_document_line" => $documentLine,
                        ]
                    )
                );
            } else {
                if ($documentLine['line_type'] === 'record') {
                    (new PrintRecordService())->ifTabTypeRecord(
                        PrintWordRecordDTO::createFromArray(
                            [
                                "key_of_current_document_line" => $key,
                                "section" => $this->section,
                                "page_width_in_twip_without_margins" => $this->pageWidthInTwipWithoutMargins,
                                "document_lines" => $documentLine,
                            ]
                        ),
                        $table
                    );
                }
            }
        }
        return $phpWord;
    }


}
