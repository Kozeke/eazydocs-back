<?php

namespace App\Http\Services\PrintWordDocument;

use App\DTO\PrintWordTableDTO;
use App\Extensions\PrintWordDocument\PrintWordTrait;
use Illuminate\Support\Collection as SupportCollection;

/**
 * Class PrintTableService
 * @package App\Http\Services\PrintWordDocument
 *
 * @author Kozy-Korpesh Tolep
 */
class PrintTableService
{
    use PrintWordTrait;

    /**
     * @var int
     */
    private const ONE_LINE_FOR_TEXT_BREAK = 1;

    /**
     * @var string
     */
    private const TABLE_STYLE_NAME_FOR_ACTS = 'Раздел 1';

    /**
     * @param PrintWordTableDTO $actEditorPrintWordTableDTO
     * @return void
     */
    public function printTable(PrintWordTableDTO $actEditorPrintWordTableDTO): void
    {
        $this->massAssignmentOfPropertiesFromDTO($actEditorPrintWordTableDTO);
        $this->printAsTable($this->currentDocumentLine);

        $this->section->addTextBreak(
            self::ONE_LINE_FOR_TEXT_BREAK,
            $this->predefinedStyles['sizeBreak'],
            $this->predefinedStyles['spaceBreak']
        );
    }

    /**
     * @param array $actEditorDocument
     * @return void
     */
    private function printAsTable(array $actEditorDocument): void
    {
        ['columns' => $columns, 'is_custom_numeration' => $isCustomNumeration] = $actEditorDocument;
        $filteredColumns = array_filter(
            $columns,
            fn($column) => $column['print'] === 'printOn'
        );
        if (count($filteredColumns) !== 0) {
            $this->section->addText(
                htmlspecialchars($actEditorDocument['fieldName']),
                [
                    'size' => $actEditorDocument['fontSize'],
                    'bold' => $actEditorDocument['titleBoldOnPrint'],
                    'italic' => $actEditorDocument['titleItalicOnPrint'],
                ],
                [
                    'align' => $actEditorDocument['titleHorizontalAlignmentOnPrint'],
                    'spaceAfter' => 0,
                ]
            );
            $this->table = $this->section->addTable(self::TABLE_STYLE_NAME_FOR_ACTS);
            $this->table->addRow();
            $this->printHeader($filteredColumns, $actEditorDocument);
            $this->printNumerationRow($filteredColumns, $actEditorDocument, $isCustomNumeration);
            $this->printBody($filteredColumns, $actEditorDocument);
        }
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
     * @param SupportCollection $rows
     * @param array $columns
     * @param array $actEditorDocument
     * @return void
     */
    private function printEmptyRow(SupportCollection $rows, array $columns, array $actEditorDocument): void
    {
        if ($rows->count() === 0) {
            $this->table->addRow();
            $this->printRow($columns, $actEditorDocument, null);
        }
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
     * @param bool $isCustomNumeration
     * @return void
     */
    private function printNumerationRow(array $columns, array $actEditorDocument, bool $isCustomNumeration): void
    {
        if ($isCustomNumeration) {
            $this->table->addRow();
            $this->printRow($columns, $actEditorDocument, 'custom_column_number');
        }
    }

    /**
     * @param array $columns
     * @param array $actEditorDocument
     * @return void
     */
    private function printBody(array $columns, array $actEditorDocument): void
    {
        $rows = $this->getRows($columns);
        $rows->each(function ($item) use ($actEditorDocument) {
            $this->table->addRow();
            $this->printRow($item->get('columns'), $actEditorDocument, 'valueForTable');
        });
        $this->printEmptyRow($rows, $columns, $actEditorDocument);
    }
}
