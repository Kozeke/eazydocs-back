<?php

namespace App\DTO;


use Illuminate\Support\Collection;
use PhpOffice\PhpWord\Element\Section;

/**
 * Class ActEditorPrintWordTableDTO
 * @package App\DTO\ActEditorPrint
 *
 * @author Kozy-Korpesh Tolep
 */
class PrintWordRecordDTO extends DTO
{

    /**
     * @var float
     */
    public float $page_width_in_twip_without_margins;

    /**
     * @var \PhpOffice\PhpWord\Element\Section
     */
    public Section $section;

    /**
     * @var \Illuminate\Support\Collection|null
     */
    public ?Collection $document_lines;

    /**
     * @var int|null
     */
    public ?int $key_of_current_document_line;


}
