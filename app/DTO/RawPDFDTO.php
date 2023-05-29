<?php

namespace App\DTO;


class RawPDFDTO extends DTO
{
    /**
     * @var array|null
     */
    public ?array $line_type;

    /**
     * @var int|null
     */
    public ?int $order;

    /**
     * @var bool|null
     */
    public ?bool $enabled;
}
