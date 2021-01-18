<?php

namespace Uccello\DocumentDesignerCore\Support;

use Illuminate\Support\Facades\Storage;

class DocumentIO
{
    public static function process(string $templateFile, string $outFile, array $data, array $options = null)
    {
        if (static::endsWith($templateFile, '.xlsx') && static::endsWith($outFile, '.xlsx')) {
            $calc = new CalcProcessor($templateFile);
            $calc->process($data);
            $calc->saveAs($outFile);
        } elseif (static::endsWith($templateFile, '.docx') && static::endsWith($outFile, '.docx')) {
            $document = new DocumentProcessor($templateFile);
            $document->processRecursive($data);
            $document->saveAs($outFile);
        } elseif (static::endsWith($templateFile, '.docx') && static::endsWith($outFile, '.pdf')) {
            Storage::makeDirectory('temp');

            $tempFileDocx = storage_path("app/temp/") . basename($outFile, '.pdf') . '.docx';
            $tempFilePdf = storage_path("app/temp/") . basename($outFile);

            $document = new DocumentProcessor($templateFile);
            $document->processRecursive($data);
            $document->saveAs($tempFileDocx);

            static::convertToPdf($tempFileDocx);

            // Storage::disk('local')->move($tempFilePdf, $outFile);
            // Storage::disk('local')->delete($tempFileDocx);
            rename($tempFilePdf, $outFile);
            unlink($tempFileDocx);
        } elseif (static::endsWith($templateFile, '.pdf') && static::endsWith($outFile, '.pdf')) {
            $document = new PdfProcessor($templateFile, $options);
            $document->process($data);
            $document->saveAs($outFile);
        } else {
            return false;
        }

        return true;
    }

    public static function getVariables(string $templateFile, array $options = null)
    {
        if (static::endsWith($templateFile, '.xlsx')) {
            $calc = new CalcProcessor($templateFile);
            $variables = $calc->getVariables($options);
        } elseif (static::endsWith($templateFile, '.docx')) {
            $document = new DocumentProcessor($templateFile);
            $variables = $document->getVariables();
        } elseif (static::endsWith($templateFile, '.pdf')) {            // TODO...
            // $document = new PdfProcessor($templateFile, $options);
            // $variables = $document->getVariables($options);
        } else {
            $variables = null;
        }

        return $variables;
    }

    protected static function endsWith($string, $endString)
    {
        $len = strlen($endString);
        if ($len == 0) {
            return true;
        }
        return (substr($string, -$len) === $endString);
    }

    protected static function convertToPdf($inFile)
    {
        $file = basename($inFile);
        $path = dirname($inFile);

        exec("/usr/bin/soffice --headless --convert-to pdf --outdir \"$path\" \"$path/$file\"");
    }
}
