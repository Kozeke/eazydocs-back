<?php

namespace App\Http\Requests;

use Illuminate\Foundation\Http\FormRequest;
use App\Rules\SettingsBasedOnLineType;

class PdfPrintRequest extends FormRequest
{
    /**
     * Determine if the user is authorized to make this request.
     *
     * @return bool
     */
    public function authorize()
    {
        return true;
    }

    /**
     * Get the validation rules that apply to the request.
     *
     * @return array<string, mixed>
     */
    public function rules()
    {
        return [
            'documentDefaultSettings' => 'filled',
            'documentDefaultSettings.leftIndentInCm' => 'nullable|numeric',
            'documentDefaultSettings.rightIndentInCm' => 'nullable|numeric',
            'documentDefaultSettings.topIndentInCm' => 'nullable|numeric',
            'documentDefaultSettings.bottomIndentInCm' => 'nullable|numeric',
            'documentDefaultSettings.fontSize' => 'nullable|integer',
            'documentDefaultSettings.substringFontSize' => 'nullable|integer',
            'documentDefaultSettings.spacing' => 'nullable|numeric',
            'documentDefaultSettings.fontFamily' => 'nullable|string',

            'documentLines' => 'filled',
            'documentLines.*.line_type' => 'required|string|in:table,record',
            'documentLines.*.order' => 'required|integer',
            'documentLines.*.settings' => 'filled',
            'documentLines.*.settings.column_size' => 'sometimes|array',
            'documentLines.*.data' => 'sometimes|array',
            'documentLines.*.title' => 'sometimes|string',
            'documentLines.*.field' => 'sometimes|string',
            'documentLines.*.substring' => 'sometimes|nullable|string',
            'documentLines.*.settings.title_settings' => 'sometimes|array|filled',
            'documentLines.*.settings.title_settings.font' => 'sometimes|integer',
            'documentLines.*.settings.title_settings.alignment' => 'sometimes|string|in:L,R,C',
            'documentLines.*.settings.title_settings.size_type' => 'sometimes|string|in:content,percent,endLine',
            'documentLines.*.settings.title_settings.percent_size' => 'sometimes|nullable|integer',
            'documentLines.*.settings.title_settings.text_type' => 'sometimes|nullable|string|in:bold,italics,underline',
            'documentLines.*.settings.field_settings.font' => 'sometimes|integer',
            'documentLines.*.settings.field_settings.alignment' => 'sometimes|string|in:L,R,C',
            'documentLines.*.settings.field_settings.size_type' => 'sometimes|string|in:content,percent,endLine',
            'documentLines.*.settings.field_settings.percent_size' => 'sometimes|nullable|integer',
            'documentLines.*.settings.field_settings.text_type' => 'sometimes|nullable|string|in:bold,italics,underline',
        ];
    }
}
