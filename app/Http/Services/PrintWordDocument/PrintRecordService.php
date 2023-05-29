<?php

namespace App\Http\Services\PrintWordDocument;
use App\DTO\PrintWordRecordDTO;
use App\Extensions\PrintWordDocument\PrintWordTrait;
use Illuminate\Support\Collection as SupportCollection;
use PhpOffice\PhpWord\Element\Table;

/**
 * Class PrintRecordService
 * @package App\Http\Services\PrintWordDocument
 *
 * @author Kozy-Korpesh Tolep
 */
class PrintRecordService
{

    use PrintWordTrait;

    /**
     * @var mixed|null
     */
    private ?SupportCollection $documentLines;

    /**
     * @var int|null
     */
    private ?int $keyOfCurrentDocumentLine;

    /**
     * @var int
     */
    private const ONE_LINE_FOR_TEXT_BREAK = 1;

    /**
     * @param PrintWordRecordDTO $PrintWordRecordDTO
     * @param Table|null $table
     * @return void
     */
    public function ifTabTypeRecord(
        PrintWordRecordDTO $PrintWordRecordDTO,
        ?Table &$table = null
    ): void {
        $this->massAssignmentOfPropertiesFromDTO($PrintWordRecordDTO);
        $this->table = &$table;
        $this->isRecordType = true;
        $documentLine = $this->$documentLines[$this->keyOfCurrentDocumentLine];
        $this->fontSize = $Document['fontSize'] ?? (int)$this->predefinedStyles['fontSize'];
        //не добавляем новую строку если предыдущая запись имеет тот же lineNumberOnPrint
        if ($this->keyOfCurrentDocumentLine === 0 || $this->isPreviousDocumentHasSameLineNumber(
                $documentLine
            )) {
            $this->countSubCell = null;
            $this->section->addTextBreak(
                self::ONE_LINE_FOR_TEXT_BREAK,
                $this->predefinedStyles['sizeBreak'],
                $this->predefinedStyles['spaceBreak']
            );
            $this->table = $this->section->addTable();
            $this->table->addRow();
        }
        $this->countSubCell = is_null($this->countSubCell) ? 0 : $this->countSubCell + 1;
        $isNextDocSameLineNumber = $this->isNextDocumentHasSameLineNumber($documentLine);
        if (
            $this->isSubCellNecessary($documentLine)
        ) {
            $this->countSubCell++;
        }
        $this->printRecord($documentLine);
        $this->cellsOnOneLine[] = $documentLine;
        if (
            ($this->isCurrentDocumentIsLast() || !$isNextDocSameLineNumber)
            && $this->needSubCell
        ) {
            $this->setCountSubCell();
            $this->createSubscriptCells();
        }
        $this->cellsOnOneLine = $isNextDocSameLineNumber ? $this->cellsOnOneLine : [];
        $this->isRecordType = false;
    }

    /**
     * @param array $actEditorDocument
     * @return bool
     */
    private function isNextDocumentHasSameLineNumber(array $actEditorDocument): bool
    {
        return ($this->actEditorDocuments[$this->keyOfCurrentDocumentLine + 1]['lineNumberOnPrint'] ?? '') === $actEditorDocument['lineNumberOnPrint'];
    }

    /**
     * @return bool
     */
    private function isCurrentDocumentIsLast(): bool
    {
        return count($this->documentLines) - 1 === $this->keyOfCurrentDocumentLine;
    }

    /**
     * @param array $actEditorDocument
     * @return bool
     */
    private function isPreviousDocumentHasSameLineNumber(array $actEditorDocument): bool
    {
        $keyOfPreviousActEditorDocument = $this->keyOfCurrentDocumentLine !== 0 ? $this->keyOfCurrentDocumentLine - 1 : '';
        return ($this->actEditorDocuments[$keyOfPreviousActEditorDocument]['lineNumberOnPrint'] ?? '') !== $actEditorDocument['lineNumberOnPrint'];
    }
}
