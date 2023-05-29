<?php

namespace App\DTO;

use Illuminate\Http\Request;
use Spatie\DataTransferObject\Arr;
use Spatie\DataTransferObject\Attributes\MapFrom;
use Spatie\DataTransferObject\DataTransferObject;
use Spatie\DataTransferObject\Exceptions\UnknownProperties;

class DTO extends DataTransferObject
{
    /**
     * @var array
     */
    protected array $items = [];

    /**
     * @return array
     */
    /**
     * @param Request $request
     * @return static
     * @throws UnknownProperties
     */
    public static function createFromRequest(Request $request): static
    {
        $dto = new static(
            $request->all()
        );
        if ($request->keys()) {
            return $dto->only(...$request->keys());
        } else {
            return $dto->except(...array_keys($dto->all()));
        }
    }

    /**
     * @param array $data
     * @return static
     * @throws UnknownProperties
     */
    public static function createFromArray(array $data): static
    {
        $keys = array_map('snake_case', array_keys($data));
        $data = array_combine($keys, $data);
        $dto = new static(
            $data
        );
        if (array_keys($data)) {
            return $dto->only(...array_keys($data));
        }
        return $dto->except(...array_keys($dto->all()));
    }

}
