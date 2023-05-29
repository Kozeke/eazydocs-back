<?php

namespace App\Http\Services\PrintPDFDocument;

use App\DTO\PrintPDFRecordDTO;
use App\Extensions\PrintWordDocument\PrintWordTrait;
use Mpdf\Mpdf;
use Mpdf\MpdfException;

/**
 * Class PrintRecordService
 * @package App\Http\Services\PrintPDFDocument
 *
 * @author Kozy-Korpesh Tolep
 */
class PrintRecordService
{
    use PrintWordTrait;

    /**
     * @var  Mpdf
     */
    private Mpdf $pdf;

    /**
     * @var int|float
     */
    private int|float $pageWidthInMmWithoutMargins;

    /**
     * @var int|null
     */
    private ?int $keyOfCurrentDocumentLine;

    /**
     * @var int
     */
    private int $titleCellWidth;

    /**
     * @var array
     */
    private array $widthOfCellsInOneLine = [];

    /**
     * @var bool
     */
    private bool $saveWidth = false;

    /**
     * @var array
     */
    private array $titleSettings;

    /**
     * @var array
     */
    private array $fieldSettings;

    /**
     * @var array
     */
    private array $substringCellsInOneLine;

    /**
     * @var int
     */
    private const CELL_HEIGHT = 4;

    /**
     * @var float
     */
    const CONSTANT_K = 6;

    /**
     * @param PrintPDFRecordDTO $printPDFRecordDTO
     * @return void
     * @throws MpdfException
     */
    public function ifTabTypeRecord(PrintPDFRecordDTO $printPDFRecordDTO): void
    {
        $this->massAssignmentOfPropertiesFromDTO($printPDFRecordDTO);
        //не добавляем новую строку если предыдущая запись имеет тот же order
        $keyOfPreviousDocumentLine = $this->keyOfCurrentDocumentLine !== 0 ? $this->keyOfCurrentDocumentLine - 1 : '';
        $documentLine = $this->documentLines[$this->keyOfCurrentDocumentLine];
        if ($this->keyOfCurrentDocumentLine === 0 || ($documentLines[$keyOfPreviousDocumentLine]['lineNumberOnPrint'] ?? '') !== $documentLine['lineNumberOnPrint']) {
            $this->pdf->Ln($this->predefinedStyles['spacing']);
        }
        $isNextActEditorDocSameLineNumber =
            (
                $documentLines[$this->keyOfCurrentDocumentLine + 1]['lineNumberOnPrint'] ?? ''
            ) === $documentLine['lineNumberOnPrint'];
        $nextdocumentLine = $documentLines[$this->keyOfCurrentDocumentLine + 1] ?? null;
        $isNextActEditorDocNotSignature =
            !$nextdocumentLine || $nextdocumentLine['tab_type'] !== 'signature';
        if (
            $documentLine['fieldTypeSizeOnPrint'] !== 'string'
            && $isNextActEditorDocNotSignature
            && $isNextActEditorDocSameLineNumber
        ) {
            $this->saveWidth = true;
        } elseif (!$isNextActEditorDocSameLineNumber) {
            $this->saveWidth = false;
            $this->titleCellWidth = 0;
        }
        $this->printRecord($documentLine);
        if ($nextdocumentLine && $nextdocumentLine['tab_type'] !== 'record') {
            $this->saveWidth = false;
            $this->substringCellsInOneLine = [];
        }
        if (!$isNextActEditorDocSameLineNumber) {
            $this->substringCellsInOneLine = [];
        }
    }

    /**
     * @param array $documentLine
     * @return void
     * @throws MpdfException
     */
    private
    function printRecord(
        array $documentLine
    ): void {
        $this->viewFieldOnPrint = $documentLine['viewFieldOnPrint'];
        $styles = $this->getPrintStyle($documentLine);
        if ($documentLine['viewTitleOnPrint']) {
            $this->titleCellWidth = $this->calculateCellWidth(
                $documentLine['titleTypeSizeOnPrint'],
                $documentLine['titleProcentSizeOnPrint'] ?? null,
                $documentLine['fieldName']
            );
            $this->printTitle($documentLine, $styles);
        } else {
            $this->titleCellWidth = 0;
        }
        if ($documentLine['viewFieldOnPrint']) {
            $this->printField($documentLine, $styles);
        }
    }

    /**
     * @param array $documentLine
     * @param array $styles
     * @throws MpdfException
     */
    private
    function printTitle(
        array $documentLine,
        array $styles
    ) {
        $documentLine['fieldName'] = is_null(
            $documentLine['fieldName']
        ) ? '' : $documentLine['fieldName'];
        $this->titleTypeSizeOnPrint = $documentLine['titleTypeSizeOnPrint'];
        switch ($documentLine['titleTypeSizeOnPrint']) {
            case 'endLine':
                $this->printWhenTypeSizeEndLine($documentLine['fieldName'], true, $styles['titleFontStyle'], null);
                break;
            case 'procent':
                $this->printWhenTypeSizeProcent($documentLine['fieldName'], true, $styles['titleFontStyle']);
                break;
            case 'content':
                $this->printWhenTypeSizeContent($styles['titleFontStyle'], $documentLine['fieldName']);
                break;
        }
    }

    /**
     * @param string $textValue
     * @param bool $isTitle
     * @param array $styles
     * @param float|null $cellWidth
     * @param string|null $subsText
     * @throws MpdfException
     */
    private
    function printWhenTypeSizeEndLine(
        string $textValue,
        bool $isTitle,
        array $styles,
        ?float $cellWidth = 0,
        ?string $subsText = ''
    ) {
        if ($isTitle) {
            $this->createCell($this->titleCellWidth, $styles, $textValue);
            if ($this->viewFieldOnPrint) {
                $this->pdf->Ln($this->predefinedStyles['spacing']);
            }
        } else {
            $this->createCell($cellWidth, $styles, $textValue, $subsText);
        }
    }

    /**
     * @param string $textValue
     * @param bool $isTitle
     * @param array $styles
     * @param float|null $cellWidth
     * @param string|null $subsText
     * @return void
     * @throws MpdfException
     */
    private
    function printWhenTypeSizeProcent(
        string $textValue,
        bool $isTitle,
        array $styles,
        ?float $cellWidth = 0,
        ?string $subsText = ''
    ): void {
        if ($isTitle) {
            $this->createCell($this->titleCellWidth, $styles, $textValue);
        } else {
            $this->createCell($cellWidth, $styles, $textValue, $subsText);
        }
    }

    /**
     * @param array $styles
     * @param string|null $textValue
     * @param float|null $cellWidth
     * @param string|null $subsText
     * @return void
     * @throws MpdfException
     */
    private
    function printWhenTypeSizeContent(
        array $styles,
        ?string $textValue,
        ?float $cellWidth = 0.0,
        ?string $subsText = ''
    ): void {
        if ($textValue) {
            $this->createCell($this->titleCellWidth, $styles, $textValue);
        } else {
            $this->createCell($cellWidth, $styles, $textValue, $subsText);
        }
    }


    /**
     * @param array $documentLine
     * @param array $styles
     * @return void
     * @throws MpdfException
     */
    private
    function printField(
        array $documentLine,
        array $styles
    ): void {
        $documentLine['textValue'] = '';
        $cellWidth = $this->calculateCellWidth(
            $documentLine['fieldTypeSizeOnPrint'],
            $documentLine['fieldProcentSizeOnPrint'],
            $documentLine['textValue']
        );
        $this->fieldTypeSizeOnPrint = $documentLine['fieldTypeSizeOnPrint'];
        switch ($documentLine['fieldTypeSizeOnPrint']) {
            case 'procent':
                $this->printWhenTypeSizeProcent(
                    $documentLine['textValue'],
                    false,
                    $styles['fieldFontStyle'],
                    $cellWidth,
                    $documentLine['subscriptOnPrint']
                );
                break;
            case 'endLine':
                $this->printWhenTypeSizeEndLine(
                    $documentLine['textValue'],
                    false,
                    $styles['fieldFontStyle'],
                    $this->pageWidthInMmWithoutMargins - $this->titleCellWidth,
                    $documentLine['subscriptOnPrint']
                );
                break;
            case 'content':
                $this->printWhenTypeSizeContent(
                    $styles['fieldFontStyle'],
                    $documentLine['textValue'],
                    $this->pageWidthInMmWithoutMargins - $this->titleCellWidth,
                    $documentLine['subscriptOnPrint']
                );
                break;
            case 'string':
                $this->printFieldWithSizeTypeString($styles['fieldFontStyle'], $documentLine);
                break;
        }
    }

    /**
     * @param array $styles
     * @param array $documentLine
     * @return void
     * @throws MpdfException
     */
    private
    function printFieldWithSizeTypeString(
        array $styles,
        array $documentLine
    ): void {
        $lines = [];
        if (!empty($documentLine['textValue'])) {
            $lines = $this->divideTextIntoLinesByPaperWidth(
                $documentLine['textValue'],
            );
        }
        if ($documentLine['titleTypeSizeOnPrint'] !== 'endLine') {
            $this->pdf->Ln($this->predefinedStyles['spacing']);
        } else {
            $cellWidth = $this->pageWidthInMmWithoutMargins - $this->titleCellWidth;
        }
        if (!empty($lines)) {
            //if the field is not empty
            if ($documentLine['fieldStringSizeOnPrint'] === 'content') {
                //calculating number of lines based on the length of the field
                foreach ($lines as $key => $line) {
                    $substring = $key === 0 ? $documentLine['subscriptOnPrint'] : '';
                    $cellWidth = isset($cellWidth) && $key === 0 ? $cellWidth : $this->pageWidthInMmWithoutMargins;
                    if ($key > 0) {
                        $this->pdf->Ln($this->predefinedStyles['spacing']);
                        if ($this->titleHorizontalAlignmentOnPrint === 'interlinearInColumnEnum') {
                            $this->pdf->Cell(
                                $this->titleCellWidth,
                                self::CELL_HEIGHT,
                                '',
                                '',
                                '',
                                $styles['alignment']
                            );
                            $cellWidth = $this->pageWidthInMmWithoutMargins - $this->titleCellWidth;
                        }
                    }
                    $this->createCell(
                        $cellWidth,
                        $styles,
                        $line,
                        $substring
                    );
                }
            } else {
                //if the number of lines are given
                foreach ($lines as $key => $line) {
                    if ($key > 0) {
                        $this->pdf->Ln($this->predefinedStyles['spacing']);
                    }
                    $this->createCell(
                        $this->pageWidthInMmWithoutMargins,
                        $styles,
                        $line,
                        $documentLine['subscriptOnPrint']
                    );
                }
                if (count($lines) < $documentLine['fieldStringSizeOnPrint']) {
                    for ($i = count($lines); $i <= $documentLine['fieldStringSizeOnPrint'] - 1; $i++) {
                        if ($i > 0) {
                            $this->pdf->Ln($this->predefinedStyles['spacing']);
                        }
                        $this->createCell(
                            $this->pageWidthInMmWithoutMargins,
                            $styles,
                            '',
                            $documentLine['subscriptOnPrint']
                        );
                    }
                }
            }
        } else {
            //if the field is empty and number of lines are given
            $documentLine['fieldStringSizeOnPrint'] = $documentLine['fieldStringSizeOnPrint'] == 'content' ? 1 : $documentLine['fieldStringSizeOnPrint'];
            for ($i = 0; $i < $documentLine['fieldStringSizeOnPrint']; $i++) {
                if ($i > 0) {
                    $this->pdf->Ln($this->predefinedStyles['spacing']);
                }
                $this->createCell(
                    $this->pageWidthInMmWithoutMargins,
                    $styles,
                    '',
                    $documentLine[documentLine::FIELDS_TYPE_SUBSCRIPT[$i]]
                );
            }
        }
    }


    /**
     * @param string $words
     * @param int|null $titleLastLineLength
     * @return array
     */
    private
    function divideTextIntoLinesByPaperWidth(
        string $words,
        ?int $titleLastLineLength = 0
    ): array {
        $line = '';
        $wordsArray = explode(' ', $words);
        $lines = [];
        foreach ($wordsArray as $key => $word) {
            if ($titleLastLineLength + $this->pdf->GetStringWidth($line) + $this->pdf->GetStringWidth(
                    $word
                ) < $this->pageWidthInMmWithoutMargins - 3) {
                $line .= " " . $word;
            } else {
                $titleLastLineLength = 0;
                $lines[] = $line;
                $line = $word;
            }
            // сумма символов слов последней линии может быть меньше или равно ширине документа
            // и в этом случае мы просто добавляем последнюю линию в массив
            if ($key === count($wordsArray) - 1) {
                $lines[] = $line;
            }
        }
        return $lines;
    }

    /**
     * @param string|null $typeSize
     * @param string|null $percent
     * @param string|null $text
     * @return float
     */
    private
    function calculateCellWidth(
        ?string $typeSize,
        ?string $percent,
        ?string $text = ''
    ): float {
        $lengthOfField = 0;
        switch ($typeSize) {
            case 'endLine':
                $lengthOfField = $this->pageWidthInMmWithoutMargins;
                break;
            case 'procent':
                $lengthOfField = ((int)$percent * $this->pageWidthInMmWithoutMargins) / 100;
                if ($text && $lengthOfField < $lengthOfText = $this->pdf->GetStringWidth($text)) {
                    return $lengthOfText;
                }
                break;
            case 'content':
                $lengthOfField = $this->pdf->GetStringWidth($text) + self::CONSTANT_K;
                break;
        }
        return $lengthOfField;
    }

    /**
     * @param float $cellWidth
     * @param array $styles
     * @param string|null $textValue
     * @param string|null $subsText
     * @throws MpdfException
     */
    private
    function createCell(
        float $cellWidth,
        array $styles,
        ?string $textValue,
        ?string $subsText = ''
    ) {
        $textValue = is_null($textValue) ?: str_replace("\n", "", $textValue);
        $subsText = is_null($subsText) ?: str_replace("\n", "", $subsText);
        $this->pdf->setFont('freeserif', $styles['style'], $styles['fontSize']);
        //decrease the font of the text of the field if it's length is larger than the cell width
        if ($this->pdf->GetStringWidth($textValue) > $cellWidth && $styles['underline'] === 'B') {
            $this->pdf->setFont('freeserif', $styles['style'], $styles['fontSize'] - 1);
        } else {
            $this->pdf->setFont('freeserif', $styles['style'], $styles['fontSize']);
        }
        //printing cell with value and comparing its length with pageWidthInMmWithoutMargins
        if ($this->pdf->GetStringWidth($textValue) < $this->pageWidthInMmWithoutMargins) {
            $this->pdf->Cell(
                $cellWidth,
                self::CELL_HEIGHT,
                ltrim($textValue),
                $styles['underline'],
                0,
                $styles['alignment']
            );
        } else {
            $lines = $this->divideTextIntoLinesByPaperWidth($textValue);
            foreach ($lines as $key => $line) {
                if ($key > 0) {
                    $this->pdf->Ln($this->predefinedStyles['spacing']);
                }
                $this->pdf->Cell(
                    $cellWidth,
                    self::CELL_HEIGHT,
                    ltrim($line),
                    $styles['underline'],
                    '',
                    $styles['alignment']
                );
            }
        }
        if ($this->saveWidth) {
            //если есть другой act_editor_tab в этой линии сохраняем подстрочник в массив так как если напечатать подстрчник с новой строки,
            //нельзя будет вернуться в предыдущую строку
            $this->substringCellsInOneLine[] = [
                'size' => $cellWidth,
                'subsText' => $subsText,
            ];
        } else {
            $addedNewLine = false;
            if ((count($this->substringCellsInOneLine) > 0 && $this->checkIfSubstringExists()) || (count(
                        $this->substringCellsInOneLine
                    ) > 0 && $subsText)) {
                //если нет другой act_editor_tab в этой линии проверяем массив substringCellsInOneLine и если в них есть подстрочник, начинаем печатать
                //или если есть подстрочник у нынешней act_editor_tab тоже начинаем печатать
                $this->pdf->Ln();
                $addedNewLine = true;
                foreach ($this->substringCellsInOneLine as $substring) {
                    $this->createSubstringCell(
                        $substring['size'],
                        null,
                        $substring['subsText']
                    );
                }
                $this->substringCellsInOneLine = [];
                $this->titleCellWidth = null;
            }
            if ($subsText) {
                if (!$addedNewLine) {
                    $this->pdf->Ln();
                }
                $this->createSubstringCell($cellWidth, $this->titleCellWidth, $subsText);
            }
        }
    }

    /**
     * @param float $cellWidth
     * @param float|null $titleCellWidth
     * @param string|null $substringText
     * @return void
     * @throws MpdfException
     */
    private
    function createSubstringCell(
        float $cellWidth,
        ?float $titleCellWidth,
        ?string $substringText = ''
    ): void {
        $this->pdf->setFont('freeserif', '', $this->predefinedStyles['substringFontSize']);
        if (
            (
                $titleCellWidth
                && $this->titleTypeSizeOnPrint
                && $this->titleTypeSizeOnPrint !== 'endLine'
                && $this->fieldTypeSizeOnPrint
                && $this->fieldTypeSizeOnPrint !== 'string'
            ) || (
                $this->titleHorizontalAlignmentOnPrint === 'interlinearInColumnEnum' && $this->viewTitleOnPrint
            ) || ($this->titleHorizontalAlignmentOnPrint === 'interlinearEnum' && $this->viewTitleOnPrint)
        ) {
            $this->pdf->Cell($this->titleCellWidth, self::CELL_HEIGHT, '');
        }
        if ($this->pdf->GetStringWidth($substringText) < $this->pageWidthInMmWithoutMargins) {
            $this->pdf->Cell($cellWidth, self::CELL_HEIGHT, $substringText, '', '', 'C');
        } elseif (!empty($substringText)) {
            $lines = $this->divideTextIntoLinesByPaperWidth($substringText);
            foreach ($lines as $key => $line) {
                if ($key > 0) {
                    $this->pdf->Ln();
                }
                $this->pdf->Cell($cellWidth, self::CELL_HEIGHT, ltrim($line), '', '', 'C');
            }
        }
        $this->pdf->setFont('freeserif', '', $this->predefinedStyles['fontSize']);
    }

}
