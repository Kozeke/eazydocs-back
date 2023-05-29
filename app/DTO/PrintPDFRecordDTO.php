<?php

namespace App\DTO;


use Illuminate\Support\Collection;

/**
 * Class ActEditorPrintWordTableDTO
 * @package App\DTO\ActEditorPrint
 *
 * @author Kozy-Korpesh Tolep
 */
class PrintPDFRecordDTO extends DTO
{

    /**
     * @var float
     */
    public float $page_width_in_twip_without_margins;

    /**
     * @var \Mpdf\Mpdf
     */
    public \Mpdf\Mpdf $pdf;

    /**
     * @var \Illuminate\Support\Collection|null
     */
    public ?Collection $document_lines;

    /**
     * @var int|null
     */
    public ?int $key_of_current_document_line;


}
