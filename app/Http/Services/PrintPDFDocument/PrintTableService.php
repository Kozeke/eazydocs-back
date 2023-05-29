<?php

namespace App\Http\Services\PrintPDFDocument;

use App\DTO\PrintPDFRecordDTO;
use App\Extensions\PrintWordDocument\PrintWordTrait;
use Mpdf\Mpdf;
use Mpdf\MpdfException;

/**
 * Class PrintTableService
 * @package App\Http\Services\PrintPDFDocument
 *
 * @author Kozy-Korpesh Tolep
 */
class PrintTableService
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
    private const CELL_HEIGHT = 4;

    /**
     * @var float
     */
    const CONSTANT_K = 6;


    /**
     * @param PrintPDFRecordDTO $printPDFRecordDTO
     * @throws MpdfException
     */
    public
    function printAsTable(
        PrintPDFRecordDTO $printPDFRecordDTO
    ): void {
        $this->massAssignmentOfPropertiesFromDTO($printPDFRecordDTO);
        ['columns' => $columns] = $this->documentLines[$this->keyOfCurrentDocumentLine];
        $documentLine = $this->documentLines[$this->keyOfCurrentDocumentLine];
        $filteredColumns = array_filter(
            $columns,
            fn($column) => $column['print'] === 'printOn'
        );
        if ($documentLine['fieldName']) {
            $this->pdf->MultiCell(0, self::CELL_HEIGHT, $documentLine['fieldName']);
        }
        $html = $this->generateHTMLForTable($filteredColumns);

        $this->pdf->writeHTML($html);
    }

    /**
     * @param array $filteredColumns
     * @return string
     */
    private function generateHTMLForTable(array $filteredColumns): string
    {
        $html = $this->getHTMLForTable();
        if (!empty($filteredColumns)) {
            $columnWidthInMM = $this->pageWidthInMmWithoutMargins / count($filteredColumns);
            $columnWidthInMM = $columnWidthInMM . "mm";
            $html = $this->printHeader($filteredColumns, $html, $columnWidthInMM);
            $html = $this->printBody($filteredColumns, $html, $columnWidthInMM);
        }
        $html .= "</table>";
        return $html;
    }

    /**
     * @param array $data
     * @return string
     */
    private function generateHTMLForAttachmentTable(array $data): string
    {
        $html = $this->getHTMLForTable();
        foreach ($data as $column) {
            $columnWidthInMM = $this->pageWidthInMmWithoutMargins / count($column);
            $columnWidthInMM = $columnWidthInMM . "mm";
            $html = $this->printRow($column, $html, $columnWidthInMM);
        }
        $html .= "</table>";
        return $html;
    }

    /**
     * @param array $columns
     * @param string $html
     * @param string $columnWidthInMM
     * @return string
     */
    private
    function printHeader(
        array $columns,
        string $html,
        string $columnWidthInMM
    ): string {
        return $this->printRow($columns, $html, $columnWidthInMM, 'field_name');
    }

    /**
     * @param SupportCollection $rows
     * @param array $columns
     * @param string $html
     * @param string $columnWidthInMM
     * @return string
     */
    private
    function printEmptyRow(
        SupportCollection $rows,
        array $columns,
        string $html,
        string $columnWidthInMM
    ): string {
        if ($rows->count() === 0) {
            return $this->printRow($columns, $html, $columnWidthInMM);
        }
        return $html;
    }

    /**
     * @param array $columns
     * @param string $html
     * @param string $columnWidthInMM
     * @return string
     */
    private
    function printBody(
        array $columns,
        string $html,
        string $columnWidthInMM
    ): string {
        $rows = $this->getRows($columns);
        foreach ($rows as $item) {
            $html = $this->printRow($item->get('columns'), $html, $columnWidthInMM, 'valueForTable');
        }
        return $this->printEmptyRow($rows, $columns, $html, $columnWidthInMM);
    }

    /**
     * @return string
     */
    private function getHTMLForTable(): string
    {
        return "<table style='overflow:wrap; font-size: 9px; text-align: center' cellspacing='0'>";
    }

    /**
     * @param array $columns
     * @param string $html
     * @param string $columnWidthInMM
     * @param string|null $keyForValue
     * @return string
     */
    private
    function printRow(
        array $columns,
        string $html,
        string $columnWidthInMM,
        ?string $keyForValue = null
    ): string {
        $html .= "<tr align='left'>";
        foreach ($columns as $column) {
            $fieldName = $keyForValue ? $column[$keyForValue] : $column;
            $height = $keyForValue ? "" : "50px";
            $html .= "<td align='center' style='width: $columnWidthInMM;border: 0.5px solid black;height:$height'>$fieldName</td>";
        }
        $html .= "</tr>";
        return $html;
    }

}
