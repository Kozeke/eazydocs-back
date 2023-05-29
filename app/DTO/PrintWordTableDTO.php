<?php

namespace App\DTO;

use PhpOffice\PhpWord\Element\Section;

/**
 * Class ActEditorPrintWordTableDTO
 * @package App\DTO\ActEditorPrint
 *
 * @author Kozy-Korpesh Tolep
 */
class PrintWordTableDTO extends DTO
{

    /**
     * @var \PhpOffice\PhpWord\Element\Section
     */
    public Section $section;

    /**
     * @var float
     */
    public float $page_width_in_twip_without_margins;

    /**
     * @var int|null
     */
    public ?int $exec_doc_id;

    /**
     * @var array|null
     */
    public ?array $current_document_line;


}
