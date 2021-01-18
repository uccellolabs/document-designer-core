<?php

namespace Uccello\DocumentDesignerCore\Support;

use Illuminate\Support\Facades\Storage;

// Note: Need PDFTK to be installed... (https://doc.ubuntu-fr.org/pdftk)

class PdfProcessor
{
    protected $templateFile;
    protected $options;
    protected $tempDataFile;

    public function __construct($templateFile, $options = null)
    {
        $this->templateFile = $templateFile;
        $this->options = $options;
    }

    public function process($data)
    {
        Storage::makeDirectory('temp');

        $this->tempDataFile = storage_path("app/temp/") . uniqid() . '.xfdf';

        $content = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<xfdf xmlns=\"http://ns.adobe.com/xfdf/\" xml:space=\"preserve\">\n<fields>\n";

        foreach ($data as $field => $value) {
            $content .= "<field name=\"$field\">\n<value>$value</value>\n</field>\n";
        }

        $content .= "</fields>\n</xfdf>\n";

        file_put_contents($this->tempDataFile, $content);
    }

    public function saveAs($outFile)
    {
        if (is_array($this->options) && $this->options['flatten']) {
            $flatten = 'flatten';
        } else {
            $flatten = '';
        }

        // TODO: Test ...
        exec("pdftk $this->templateFile fill_form $this->tempDataFile output $outFile $flatten", $output);

        unlink($this->tempDataFile);
        $this->tempDataFile = null;
    }
}
