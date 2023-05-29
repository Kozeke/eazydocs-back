<?php

namespace App\Extensions\PrintWordDocument;


use Illuminate\Support\Collection;
use PhpOffice\PhpWord\Element\Section;
use PhpOffice\PhpWord\Element\Table;
use PhpOffice\PhpWord\Shared\Converter;

/**
 * Class ActEditorPrintWordTrait
 * @package App\Extensions\ActEditor\PrintWord
 *
 * @author Tolep Kozy-Korpesh
 */
trait PrintWordTrait
{

    /**
     * @var array
     */
    private array $predefinedStyles =
        [
            //Стили разделов:
            'styleSection' => [
                'marginLeft' => '',
                'marginRight' => '',
                'marginTop' => '',
                'marginBottom' => '',
                'headerHeight' => 50,
                'footerHeight' => 50,
            ],
            //Стили ячеек:
            'styleCell' => [
                'valign' => 'center',
            ],
            'styleCellLine' => [
                'valign' => 'bottom',
                'borderSize' => 2,
                'borderBottomColor' => '000000',
                'borderTopColor' => 'FFFFFF',
                'borderLeftColor' => 'FFFFFF',
                'borderRightColor' => 'FFFFFF',
            ],
            'styleCellNoBorder' => [
                'valign' => 'top',
                'borderSize' => 2,
                'borderBottomColor' => 'FFFFFF',
                'borderTopColor' => 'FFFFFF',
                'borderLeftColor' => 'FFFFFF',
                'borderRightColor' => 'FFFFFF',
            ],
            'sizeBreak' => [
                'size' => 1,
            ],
            'spaceBreak' => [
                'spaceAfter' => 1,
            ],
            'spacing' => '',
            'fontSize' => '',
            //стиль таблиц
            'styleTable' => [
                'borderSize' => 1,
                'borderColor' => '000000',
                'cellMargin' => 80,
                'align' => 'center',
            ],
            //стили подстрочников
            'styleComments' => [
                'size' => '',
                'italic' => true,
            ],
            //стиль параграфа в тексте
            'styleTextParagraph' => [
                'align' => 'center',
                'spaceAfter' => 0,
                'spaceBefore' => 0
            ],
            'styleHeaderParagraph' => [
                'align' => 'left',
                'spaceAfter' => 0,
                'spaceBefore' => 0
            ],
            //стиль таблицы для подписи
            'styleTableForSigns' => [
                'borderSize' => 1,
                'borderColor' => '000000',
                'cellMargin' => 80,
                'align' => 'left',
            ],
        ];

    /**
     * @var Section
     */
    private Section $section;

    /**
     * @var Table|null
     */
    private ?Table $table = null;

    /**
     * @var float
     */
    private float $pageWidthInTwipWithoutMargins;

    /**
     * @var array|null
     */
    private ?array $currentDocumentLine = null;

    /**
     * @var Collection|null
     */
    private ?Collection $documentLines;

    /**
     * @var string|null
     */
    private ?string $titleHorizontalAlignmentOnPrint = null;

    /**
     * @var float
     */
    private float $titleLastLineLenghtInTwip;

    /**
     * @var bool
     */
    private bool $isRecordType = false;

    /**
     * @var bool
     */
    private bool $isSignatureType = false;

    /**
     * @var int
     */
    private int $fontSize;

    /**
     * @param string|null $typeSize
     * @param string|null $percent
     * @param string|null $text
     * @param int $titleSize
     * @param bool $isTitleBold
     * @param bool $isTitleItalic
     * @return float|int|null
     */
    private function calculateLengthOfTextInTwip(
        ?string $typeSize,
        ?string $percent,
        ?string $text,
        int $titleSize = 0,
        bool $isTitleBold = false,
        bool $isTitleItalic = false
    ): float|int|null {
        $lengthOfTextInTwip = 0;
        switch ($typeSize) {
            case 'procent':
                $lengthOfTextInTwip = ($percent * $this->pageWidthInTwipWithoutMargins) / 100;
                break;
            case 'endLine':
                $lengthOfTextInTwip = $this->pageWidthInTwipWithoutMargins - $titleSize;
                break;
            case 'content':
                $symbolLengthInTwip = $this->calculateSymbolLengthAndConvertToTwip(
                    $isTitleItalic,
                    $isTitleBold
                );
                $numberOfSymbols = mb_strlen($text);
                $lengthOfTextInTwip = $numberOfSymbols * $symbolLengthInTwip;
                break;
        }
        return $lengthOfTextInTwip;
    }

    /**
     * @param bool $italic
     * @param bool $bold
     * @return float
     */
    private function calculateSymbolLengthAndConvertToTwip(
        bool $italic = false,
        bool $bold = false
    ): float {
        $italicSize = $italic ? 1.1 : 1;
        $boldSize = $bold ? 1.1 : 1;
        $symbolLengthInTwip = $this->fontSize * self::CONSTANT_K * $italicSize * $boldSize;
        return Converter::cmToTwip($symbolLengthInTwip);
    }

    /**
     * @param array $columns
     * @param array $actEditorDocument
     * @return void
     */
    private function printHeader(array $columns, array $actEditorDocument): void
    {
        $this->printRow($columns, $actEditorDocument);
    }

    /**
     * @param array $columns
     * @param array $actEditorDocument
     * @param string|null $keyForValue
     * @return void
     */
    private function printRow(
        array $columns,
        array $actEditorDocument,
        ?string $keyForValue = 'field_name'
    ): void {
        foreach ($columns as $column) {
            $columnSize = $this->calculateLengthOfTextInTwip(
                'procent',
                $column['field_procent_size_on_screen'],
                $column['field_name']
            );
            $this->table->addCell($columnSize, $this->predefinedStyles['styleCell'])->addText(
                htmlspecialchars($keyForValue ? $column[$keyForValue] : ''),
                ['size' => $actEditorDocument['fontSize']],
                ['align' => 'center', 'spaceAfter' => 0, 'spaceBefore' => 0]
            );
        }
    }

    /**
     * @param array $actEditorDocument
     * @return array
     */
    private function getPrintStyle(array $actEditorDocument): array
    {
        $titleAlign = $actEditorDocument['titleHorizontalAlignmentOnPrint'] ?? 'left';
        $fieldAlign = $actEditorDocument['tab_type'] === 'table'
            ? $actEditorDocument['columnsAlign'] ?? 'left'
            : $actEditorDocument['fieldHorizontalAlignmentOnPrint'] ?? 'left';
        $size = $actEditorDocument['fontSize'] ?? $this->predefinedStyles['fontSize'];
        $sizeForSub = $actEditorDocument['interlinear_font_size'] ?? $this->subStyles['size'];

        $actEditorDocument = $this->getActEditorDocumentWithBoldItalicForSignature($actEditorDocument);

        $titleParagraphStyle = [
            'align' => $titleAlign,
            'spacing' => $this->predefinedStyles['spacing'],
            'spaceAfter' => 0,
        ];
        $titleFontStyle = [
            'size' => $size,
            'sizeForSub' => $sizeForSub,
            'bold' => $actEditorDocument['titleBoldOnPrint'] ?? false,
            'italic' => $actEditorDocument['titleItalicOnPrint'] ?? false,
        ];
        $fieldFontStyle = [
            'size' => $size,
            'sizeForSub' => $sizeForSub,
            'bold' => $actEditorDocument['fieldBoldOnPrint'] ?? false,
            'italic' => $actEditorDocument['fieldItalicOnPrint'] ?? false,
        ];
        $tabFontStyle = [
            'size' => $size,
            'bold' => $actEditorDocument['fieldBoldOnPrint'] ?? false,
        ];
        $fieldParagraphStyle = [
            'align' => $fieldAlign,
            'spacing' => $this->predefinedStyles['spacing'],
            'spaceAfter' => 0,
        ];

        return [
            'tabFontStyle' => $tabFontStyle,
            'fieldParagraphStyle' => $fieldParagraphStyle,
            'titleParagraphStyle' => $titleParagraphStyle,
            'titleFontStyle' => $titleFontStyle,
            'fieldFontStyle' => $fieldFontStyle,
        ];
    }

    /**
     * @param array $actEditorDocument
     * @return array
     */
    private function getActEditorDocumentWithBoldItalicForSignature(array $actEditorDocument): array
    {
        if ($actEditorDocument['tab_type'] === 'signature') {
            $actEditorDocument['titleBoldOnPrint'] = $actEditorDocument['title_bold_on_print'];
            $actEditorDocument['titleItalicOnPrint'] = $actEditorDocument['title_italic_on_print'];
            $actEditorDocument['fieldBoldOnPrint'] = $actEditorDocument['field_bold_on_print'];
            $actEditorDocument['fieldItalicOnPrint'] = $actEditorDocument['field_italic_on_print'];
        }

        return $actEditorDocument;
    }

    /**
     * @param array $actEditorDocument
     * @param array $printStyle
     * @return bool|void
     */
    private function printTitle(
        array $actEditorDocument,
        array $printStyle
    ) {
        $lineBreak = false;
        $this->titleTypeSizeOnPrint = $actEditorDocument['titleTypeSizeOnPrint'];
        $actEditorDocument['fieldName'] = is_null(
            $actEditorDocument['fieldName']
        ) ? '' : $actEditorDocument['fieldName'];
        if ($actEditorDocument['viewTitleOnPrint']) {
            switch ($actEditorDocument['titleTypeSizeOnPrint']) {
                case 'procent':
                    $this->printWhenTypeSizePercentOrEndLine(
                        $actEditorDocument['fieldName'],
                        $printStyle,
                        true,
                        '',
                        $this->titleLastLineLenghtInTwip
                    );
                    break;
                case 'endLine':
                    $this->printWhenTypeSizePercentOrEndLine(
                        $actEditorDocument['fieldName'],
                        $printStyle,
                        true
                    );
                    if ($actEditorDocument['viewFieldOnPrint'] && $actEditorDocument['fieldTypeSizeOnPrint'] !== 'string') {
                        $lineBreak = true;
                    }
                    break;
                case 'content':
                    $this->printTitleWhenTypeSizeContent(
                        $actEditorDocument,
                        $this->titleLastLineLenghtInTwip,
                        $printStyle,
                    );
                    break;
            }
            return $lineBreak;
        }
    }

    /**
     * @param string $words
     * @param int|null $titleLastLineLength
     * @param int|null $max
     * @return array
     */
    private
    function divideTextIntoLinesByPaperWidth(
        string $words,
        ?int $titleLastLineLength = 0,
        ?int $max = null
    ): array {
        $line = '';
        $wordsArray = explode(' ', $words);
        $linesOfText = [];
        $lineNumberOfSymbols = ceil(
            $this->pageWidthInTwipWithoutMargins / $this->calculateSymbolLengthAndConvertToTwip()
        );
        if ($max) {
            foreach ($wordsArray as $key => $word) {
                if ($titleLastLineLength + mb_strlen($line) + mb_strlen($word) < $lineNumberOfSymbols) {
                    $line .= " $word";
                } elseif ($max !== 1 && count($linesOfText) < $max - 1) {
                    $titleLastLineLength = 0;
                    $linesOfText[] = $line;
                    $line = $word;
                } else {
                    $line .= " $word";
                }
                // сумма символов слов последней линии может быть меньше или равно ширине документа
                // и в этом случае мы просто добавляем последнюю линию в массив
                if ($key === count($wordsArray) - 1) {
                    $linesOfText[] = $line;
                }
            }
        } else {
            foreach ($wordsArray as $key => $word) {
                if ($titleLastLineLength + mb_strlen($line) + mb_strlen($word) < $lineNumberOfSymbols) {
                    $line .= " $word";
                } else {
                    $titleLastLineLength = 0;
                    $linesOfText[] = $line;
                    $line = $word;
                }
                // сумма символов слов последней линии может быть меньше или равно ширине документа
                // и в этом случае мы просто добавляем последнюю линию в массив
                if ($key === count($wordsArray) - 1) {
                    $linesOfText[] = $line;
                }
            }
        }
        return $linesOfText;
    }

    /**
     * @param array $actEditorDocument
     * @return void
     */
    public function printRecord(
        array $actEditorDocument
    ): void {
        $printStyle = $this->getPrintStyle($actEditorDocument);
        $this->titleLastLineLenghtInTwip = $this->calculateLengthOfTextInTwip(
            $actEditorDocument['titleTypeSizeOnPrint'],
            $actEditorDocument['titleProcentSizeOnPrint'] ?? null,
            $actEditorDocument['fieldName'],
        );
        if ($actEditorDocument['viewTitleOnPrint']) {
            $lineBreak = $this->printTitle($actEditorDocument, $printStyle);
            if ($lineBreak) {
                $this->section->addTextBreak(
                    self::ONE_LINE_FOR_TEXT_BREAK,
                    $this->predefinedStyles['sizeBreak'],
                    $this->predefinedStyles['spaceBreak']
                );
                $this->table = $this->section->addTable();
                $this->table->addRow();
            }
        }
        if ($actEditorDocument['viewFieldOnPrint']) {
            $this->printField($actEditorDocument, $printStyle);
        }
    }

    /**
     * @param string $text
     * @param array $printStyle
     * @param bool|null $isTitle
     * @param float|null $size
     * @param string|null $subsText
     * @return void
     */
    private function printWhenTypeSizePercentOrEndLine(
        string $text,
        array $printStyle,
        ?bool $isTitle = false,
        ?string $subsText = '',
        ?float $size = 0
    ): void {
        if ($isTitle) {
            $cellStyle = $this->predefinedStyles["styleCellNoBorder"];
            $textStyle = [
                'tabStyle' => $printStyle['titleFontStyle'],
                'paragraphStyle' => $printStyle['titleParagraphStyle'],
            ];
        } else {
            $cellStyle = $this->predefinedStyles["styleCellLine"];
            $textStyle = [
                'tabStyle' => $printStyle['fieldFontStyle'],
                'paragraphStyle' => $printStyle['fieldParagraphStyle'],
            ];
        }
        $this->createCell(
            $text,
            $size ?: $this->pageWidthInTwipWithoutMargins,
            $textStyle,
            $cellStyle,
            $subsText
        );
    }

    /**
     * @param array $actEditorDocument
     * @param float $size
     * @param array $printStyle
     * @return void
     */
    private function printTitleWhenTypeSizeContent(
        array $actEditorDocument,
        float $size,
        array $printStyle
    ): void {
        $cellStyle = $this->predefinedStyles["styleCellNoBorder"];
        $textStyle = [
            'tabStyle' => $printStyle['titleFontStyle'],
            'paragraphStyle' => $printStyle['titleParagraphStyle'],
        ];
        $this->createCell(
            $actEditorDocument['fieldName'],
            $size,
            $textStyle,
            $cellStyle
        );
    }

    /**
     * @param string $text
     * @param float $size
     * @param array $printStyle
     * @param array $cellStyle
     * @param string|null $subsText
     * @param bool|null $isFieldTypeNotString
     * @return void
     */
    private function createCell(
        string $text,
        float $size,
        array $printStyle,
        array $cellStyle,
        ?string $subsText = '',
        ?bool $isFieldTypeNotString = true
    ): void {
        $this->table->addCell($size, $cellStyle)->addText(
            htmlspecialchars($text),
            $printStyle['tabStyle'],
            $printStyle['paragraphStyle']
        );
        if ($subsText) {
            $this->subStyles['size'] = $printStyle['tabStyle']['sizeForSub'];
            if (($this->isRecordType || $this->isSignatureType) && $isFieldTypeNotString) {
                $cellData = [
                    'size' => $size,
                    'printStyle' => $printStyle,
                    'subsText' => $subsText,
                ];
                $this->subCells[$this->countSubCell] = $cellData;
                if ($this->currentActEditorDocument) {
                    $this->subscriptRecordCells[$this->currentActEditorDocument['id']] = $cellData;
                }
                $this->needSubCell = true;
            } else {
                $this->createCellForSubs($size, $printStyle, $subsText);
            }
        }
    }

    /**
     * @param float $size
     * @param array $printStyle
     * @param string|null $subsText
     * @return void
     */
    private function createCellForSubs(float $size, array $printStyle, ?string $subsText = ''): void
    {
        $this->table->addRow();
        if (
            $this->isTitleTypeSizeIsNotEndLineAndFieldTypeSizeIsNotString() || (
                $this->titleHorizontalAlignmentOnPrint === 'interlinearInColumnEnum'
            ) || ($this->titleHorizontalAlignmentOnPrint === 'interlinearEnum')
        ) {
            $this->table->addCell(
                $this->titleLastLineLenghtInTwip,
                $this->predefinedStyles["styleCellNoBorder"]
            )->addText(
                '',
                $printStyle['tabStyle'],
                $printStyle['paragraphStyle']
            );
        }
        $this->table->addCell($size, $this->predefinedStyles["styleCellNoBorder"])->addText(
            htmlspecialchars($subsText),
            $this->subStyles,
            ['align' => 'center']
        );
    }

    /**
     * @param array $actEditorDocument
     * @param array $printStyle
     * @return void
     */
    private function printField(
        array $actEditorDocument,
        array $printStyle
    ): void {
        $actEditorDocument['textValue'] = '';
        if (!is_null($this->model)) {
            $this->assignValueForTextValue($actEditorDocument);
        }
        $this->fieldTypeSizeOnPrint = $actEditorDocument['fieldTypeSizeOnPrint'];
        $fieldLengthInTwip = $this->calculateLengthOfTextInTwip(
            $actEditorDocument['fieldTypeSizeOnPrint'],
            $actEditorDocument['fieldProcentSizeOnPrint'] ?? null,
            $actEditorDocument['textValue'],
        );
        if ($actEditorDocument['viewFieldOnPrint']) {
            switch ($actEditorDocument['fieldTypeSizeOnPrint']) {
                case 'procent':
                    $this->printWhenTypeSizePercentOrEndLine(
                        $actEditorDocument['textValue'],
                        $printStyle,
                        false,
                        $actEditorDocument['subscriptOnPrint'] ?? '',
                        $fieldLengthInTwip,
                    );
                    break;
                case 'endLine':
                    $this->printWhenTypeSizePercentOrEndLine(
                        $actEditorDocument['textValue'],
                        $printStyle,
                        false,
                        $actEditorDocument['subscriptOnPrint'] ?? '',
                        $this->pageWidthInTwipWithoutMargins - $this->titleLastLineLenghtInTwip,
                    );
                    break;
                case 'content':
                    $this->createFieldWithSizeTypeContent(
                        $actEditorDocument,
                        $this->titleLastLineLenghtInTwip,
                        $fieldLengthInTwip,
                        [
                            'tabStyle' => $printStyle['fieldFontStyle'],
                            'parStyle' => $printStyle['fieldParagraphStyle'],
                        ]
                    );
                    break;
                case 'string':
                    $this->createFieldWithSizeTypeString(
                        $actEditorDocument,
                        [
                            'tabStyle' => $printStyle['fieldFontStyle'],
                            'paragraphStyle' => $printStyle['fieldParagraphStyle'],
                        ]
                    );
                    break;
            }
        }
    }


    /**
     * @param array $actEditorDocument
     * @param float $titleSize
     * @param float $fieldLengthInTwip
     * @param array $printStyle
     * @return void
     */
    private function createFieldWithSizeTypeContent(
        array $actEditorDocument,
        float $titleSize,
        float $fieldLengthInTwip,
        array $printStyle
    ): void {
        if (!$actEditorDocument['textValue']) {
            $this->createCell(
                '',
                $this->pageWidthInTwipWithoutMargins,
                array(
                    'tabStyle' => $printStyle['tabStyle'],
                    'paragraphStyle' => $printStyle['parStyle'],
                ),
                $this->predefinedStyles["styleCellLine"],
                $actEditorDocument['subscriptOnPrint']
            );
        } else {
            $line = "";
            //если тип размера названия записи - до конца строки или 100%, и если линии больше одного,
            //то последняя линия поля может отличаться длиной от других
            $fieldLines = [];
            if (!empty($actEditorDocument['textValue'])) {
                $fieldLines = $this->assignValueForFieldLine($actEditorDocument, $fieldLengthInTwip);
            }
            if (!empty($fieldLines)) {
                $this->createFieldWithSizeTypeContentWhenFieldHasValue(
                    $fieldLines,
                    $actEditorDocument,
                    $titleSize,
                    $printStyle
                );
            } else {
                $this->createFieldWithSizeTypeContentWhenFieldIsEmpty($fieldLengthInTwip, $line, $printStyle);
            }
        }
    }


    /**
     * @param array $fieldLines
     * @param array $actEditorDocument
     * @param float $titleSize
     * @param array $printStyle
     * @return void
     */
    private function createFieldWithSizeTypeContentWhenFieldHasValue(
        array $fieldLines,
        array $actEditorDocument,
        float $titleSize,
        array $printStyle
    ): void {
        foreach ($fieldLines as $key => $fieldLine) {
            reset($fieldLines);
            $fieldLengthInTwip = $this->calculateLengthOfTextInTwip(
                $actEditorDocument['fieldTypeSizeOnPrint'],
                $actEditorDocument['fieldProcentSizeOnPrint'],
                $fieldLine,
                $titleSize
            );
            if ($key === key($fieldLines)) {
                $this->table->addCell($fieldLengthInTwip, $this->predefinedStyles["styleCellLine"])->addText(
                    htmlspecialchars($fieldLine),
                    $printStyle['tabStyle'],
                    $printStyle['parStyle']
                );
                $this->section->addTextBreak(
                    self::ONE_LINE_FOR_TEXT_BREAK,
                    $this->predefinedStyles['sizeBreak'],
                    $this->predefinedStyles['spaceBreak']
                );
                $this->table = $this->section->addTable();
            } else {
                $this->table->addRow();
                $this->createCell(
                    htmlspecialchars($fieldLine),
                    $fieldLengthInTwip + self::ADDITIONAL_WIDTH_TO_TABLE_IN_TWIP,
                    array(
                        'tabStyle' => $printStyle['tabStyle'],
                        'paragraphStyle' => $printStyle['parStyle'],
                    ),
                    $this->predefinedStyles["styleCellLine"],
                    $actEditorDocument['subscriptOnPrint']
                );
            }
        }
    }

    /**
     * @param float $fieldLengthInTwip
     * @param string $line
     * @param array $printStyle
     * @return void
     */
    private function createFieldWithSizeTypeContentWhenFieldIsEmpty(
        float $fieldLengthInTwip,
        string $line,
        array $printStyle
    ): void {
        $this->table->addRow();
        $this->table->addCell(
            $fieldLengthInTwip - $this->titleLastLineLenghtInTwip,
            $this->predefinedStyles["styleCellLine"]
        )->addText(
            htmlspecialchars($line),
            $printStyle['tabStyle'],
            $printStyle['parStyle']
        );
    }

    /**
     * @param array $actEditorDocument
     * @param array $printStyle
     * @return void
     */
    private function createFieldWithSizeTypeString(
        array $actEditorDocument,
        array $printStyle
    ): void {
        $this->setFieldStringSizeOnPrint($actEditorDocument);
        $this->section->addTextBreak(
            self::ONE_LINE_FOR_TEXT_BREAK,
            $this->predefinedStyles['sizeBreak'],
            $this->predefinedStyles['spaceBreak']
        );
        $linesOfText = [];
        if (!empty($actEditorDocument['textValue'])) {
            $linesOfText = $this->divideTextIntoLinesByPaperWidth(
                $actEditorDocument['textValue'],
                self::TITLE_LAST_LINE_LENGTH,
                $actEditorDocument['fieldStringSizeOnPrint']
            );
        }
        if ($actEditorDocument['fieldStringSizeOnPrint'] === 'content') {
            $this->createFieldSizeTypeStringIfFieldStringSizeIsContent($linesOfText, $printStyle, $actEditorDocument);
        } else {
            $this->createFieldSizeTypeStringIfFieldStringSizeIsNumber($linesOfText, $printStyle, $actEditorDocument);
        }
    }

    /**
     * @param array $linesOfText
     * @param array $printStyle
     * @param array $actEditorDocument
     * @return void
     */
    private function createFieldSizeTypeStringIfFieldStringSizeIsContent(
        array $linesOfText,
        array $printStyle,
        array $actEditorDocument
    ): void {
        if (!empty($linesOfText)) {
            foreach ($linesOfText as $line) {
                $this->table->addRow();
                $this->createCell(
                    $line,
                    $this->pageWidthInTwipWithoutMargins,
                    $printStyle,
                    $this->predefinedStyles["styleCellLine"],
                    $actEditorDocument['subscriptOnPrint']
                );
            }
        } else {
            if ($actEditorDocument['viewTitleOnPrint']) {
                $this->table->addRow();
            }
            $this->createCell(
                $actEditorDocument['textValue'],
                $this->pageWidthInTwipWithoutMargins,
                $printStyle,
                $this->predefinedStyles["styleCellLine"],
                $actEditorDocument['subscriptOnPrint']
            );
        }
    }

    /**
     * @param array $linesOfText
     * @param array $printStyle
     * @param array $actEditorDocument
     * @return void
     */
    private function createFieldSizeTypeStringIfFieldStringSizeIsNumber(
        array $linesOfText,
        array $printStyle,
        array $actEditorDocument
    ) {
        if (!empty($linesOfText)) {
            $this->createFieldSizeTypeStringIfFieldStringSizeIsNumberAndLinesOfTextHasValue(
                $linesOfText,
                $printStyle,
                $actEditorDocument
            );
        } else {
            $this->createFieldSizeTypeStringIfFieldStringSizeIsNumberAndLinesOfTextIsEmpty(
                $actEditorDocument,
                $printStyle
            );
        }
    }

    /**
     * @param array $linesOfText
     * @param array $printStyle
     * @param array $actEditorDocument
     * @return void
     */
    private function createFieldSizeTypeStringIfFieldStringSizeIsNumberAndLinesOfTextHasValue(
        array $linesOfText,
        array $printStyle,
        array $actEditorDocument
    ): void {
        foreach ($linesOfText as $key => $line) {
            $this->table->addRow();
            $this->createCell(
                $line,
                $this->pageWidthInTwipWithoutMargins,
                $printStyle,
                $this->predefinedStyles["styleCellLine"],
                $actEditorDocument[ActEditorDocument::FIELDS_TYPE_SUBSCRIPT[$key] ?? null]
                ?? $actEditorDocument['subscriptOnPrint'],
                false
            );
        }
        if (count($linesOfText) < $actEditorDocument['fieldStringSizeOnPrint']) {
            for ($i = count($linesOfText); $i <= $actEditorDocument['fieldStringSizeOnPrint'] - 1; $i++) {
                $this->table->addRow();
                $this->createCell(
                    '',
                    $this->pageWidthInTwipWithoutMargins,
                    $printStyle,
                    $this->predefinedStyles["styleCellLine"],
                    $actEditorDocument[ActEditorDocument::FIELDS_TYPE_SUBSCRIPT[$i]],
                    false
                );
            }
        }
    }

    /**
     * @param array $printStyle
     * @param array $actEditorDocument
     * @return void
     */
    private function createFieldSizeTypeStringIfFieldStringSizeIsNumberAndLinesOfTextIsEmpty(
        array $actEditorDocument,
        array $printStyle
    ): void {
        $actEditorDocument['fieldStringSizeOnPrint'] = $actEditorDocument['fieldStringSizeOnPrint'] == 'content' ? 1 : $actEditorDocument['fieldStringSizeOnPrint'];
        for ($i = 0; $i < $actEditorDocument['fieldStringSizeOnPrint']; $i++) {
            if (
                (
                    $actEditorDocument['viewTitleOnPrint']
                    && $actEditorDocument['titleTypeSizeOnPrint']
                )
                || $i !== 0
            ) {
                $this->table->addRow();
            }
            $this->createCell(
                '',
                $this->pageWidthInTwipWithoutMargins,
                $printStyle,
                $this->predefinedStyles["styleCellLine"],
                $actEditorDocument[ActEditorDocument::FIELDS_TYPE_SUBSCRIPT[$i]],
                false
            );
        }
    }


    /**
     * @return void
     */
    private function setOneLineCells(): void
    {
        foreach ($this->cellsOnOneLine as $cell) {
            [
                'viewTitleOnPrint' => $viewTitleOnPrint,
                'fieldName' => $fieldName,
                'viewFieldOnPrint' => $viewFieldOnPrint,
                'titleTypeSizeOnPrint' => $titleTypeSizeOnPrint,
            ] = $cell;
            if (
                $viewTitleOnPrint
                && $fieldName
                && $viewFieldOnPrint
                && $titleTypeSizeOnPrint !== 'endLine'
            ) {
                $this->oneLineCells[] = null;
                $this->oneLineCells[] = $cell;
            } elseif ($viewTitleOnPrint || $viewFieldOnPrint) {
                $this->oneLineCells[] = $cell;
            }
        }
    }

    /**
     * @param array|null $cellsData
     */
    private function addSubscriptCell(?array $cellsData): void
    {
        if (isset($cellsData)) {
            $this->table->addCell($cellsData['size'], $this->predefinedStyles["styleCellNoBorder"])
                ->addText(
                    htmlspecialchars($cellsData['subsText']),
                    $this->subStyles,
                    ['align' => 'center']
                );
        } else {
            $this->table->addCell(
                null,
                $this->predefinedStyles["styleCellNoBorder"]
            )->addText(
                '',
            );
        }
    }

    /**
     * @param mixed $DTO
     * @return void
     */
    private function massAssignmentOfPropertiesFromDTO($DTO): void
    {
        foreach ($DTO as $key => $val) {
            $this->{camel_case($key)} = $val;
        }
    }

    /**
     * @param array $actEditorDocument
     * @param float $fieldLengthInTwip
     * @return bool
     */
    private function isSizeOfFieldLargerThanPageWidth(array $actEditorDocument, float $fieldLengthInTwip): bool
    {
        return ($actEditorDocument['titleProcentSizeOnPrint'] === 100) && $fieldLengthInTwip > $this->pageWidthInTwipWithoutMargins;
    }

    /**
     * @param array $actEditorDocument
     * @return void
     */
    private function setFieldStringSizeOnPrint(array &$actEditorDocument): void
    {
        $actEditorDocument['fieldStringSizeOnPrint'] = $actEditorDocument['fieldStringSizeOnPrint'] === 'content' ? 'content' : Util::toUInt(
            $actEditorDocument['fieldStringSizeOnPrint']
        );
    }

    /**
     * @return bool
     */
    private function isTitleTypeSizeIsNotEndLineAndFieldTypeSizeIsNotString(): bool
    {
        return (
            $this->titleTypeSizeOnPrint
            && $this->titleTypeSizeOnPrint !== 'endLine'
            && $this->fieldTypeSizeOnPrint
            && $this->fieldTypeSizeOnPrint !== 'string'
        );
    }


}
