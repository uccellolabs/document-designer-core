<?php

namespace Uccello\DocumentDesignerCore\Support;

use PhpOffice\PhpWord\TemplateProcessor;
use PhpOffice\PhpWord\Exception\Exception;

class DocumentProcessor extends TemplateProcessor
{
    public function processRecursive($data)
    {
        // collect document parts
        $searchParts = array(
            $this->getMainPartName() => &$this->tempDocumentMainPart,
        );
        foreach (array_keys($this->tempDocumentHeaders) as $headerIndex) {
            $searchParts[$this->getHeaderName($headerIndex)] = &$this->tempDocumentHeaders[$headerIndex];
        }
        foreach (array_keys($this->tempDocumentFooters) as $headerIndex) {
            $searchParts[$this->getFooterName($headerIndex)] = &$this->tempDocumentFooters[$headerIndex];
        }

        $data = $this->sortRecursiveData($data);

        foreach ($searchParts as $partFileName => &$partContent) {
            foreach ($data as $key => $value) {
                $partContent = $this->processRecursiveNode($partContent, $partFileName, $key, $value);
            }
        }
    }

    private function processRecursiveNode($xml, $partFileName, $key, $data)
    {
        if (!empty($xml)) {
            $type   = substr($key, 0, 2);

            if ($type == 's:' || $type == 't:' || $type == 'b:' || $type == 'i:') {
                $key = substr($key, 2);
            }

            if ($type == 's:') { // Sheet
                // Ignore: Only for Xlsx Templates
            } elseif ($type == 't:') { // Table
                $xml = $this->processRowForScope($xml, $partFileName, $key, $data);
            } elseif ($type == 'b:') { // Block
                $xml = $this->processBlockForScope($xml, $partFileName, $key, $data);
            } elseif ($type == 'i:') { // Image
                $xml = $this->setImageValueForScope($xml, $partFileName, $key, $data);
            } else { // Vars
                $xml = $this->setValueForScope($xml, $key, $data);
            }
        }

        return $xml;
    }

    private function processRowForScope($xml, $partFileName, $key, $data)
    {
        $xml = $this->cloneRowForScope($xml, $key, count($data));

        foreach ($data as $iLine => $vars) {
            $i = $iLine + 1;

            foreach ($vars as $search => $replace) {
                $type   = substr($search, 0, 2);

                if ($type == 'i:') { // Image
                    $search = substr($search, 2);

                    $xml = $this->setImageValueForScope($xml, $partFileName, "$search#$i", $replace);
                } else { // Vars
                    $xml = $this->setValueForScope($xml, "$search#$i", $replace);
                }
            }
        }

        return $xml;
    }

    private function processBlockForScope($xml, $partFileName, $key, $data)
    {
        $xmlBlock = null;
        preg_match(
            '/(<.*)(<w:p\b.*>\${' . $key . '}<\/w:.*?p>)(.*)(<w:p\b.*\${\/' . $key . '}<\/w:.*?p>)/is',
            $xml,
            $matches
        );

        if (isset($matches[3])) {
            $xmlBlock = $matches[3];

            $cloned = array();

            foreach ($data as $vars) {
                $xB = $xmlBlock;
                $vars = $this->sortRecursiveData($vars);

                foreach ($vars as $k => $v) {
                    $xB = $this->processRecursiveNode($xB, $partFileName, $k, $v);
                }

                $cloned[] = $xB;
            }

            $xml = str_replace(
                $matches[2] . $matches[3] . $matches[4],
                implode('', $cloned),
                $xml
            );
        }

        return $xml;
    }

    protected function sortRecursiveData($data)
    {
        $blocks = [];
        $tables = [];
        $images = [];
        $vars   = [];

        foreach ($data as $key => $value) {
            $type = substr($key, 0, 2);

            if ($type == 'b:') {
                $blocks[$key] = $value;
            }
            if ($type == 't:') {
                $tables[$key] = $value;
            }
            if ($type == 'i:') {
                $images[$key] = $value;
            } else {
                $vars[$key] = $value;
            }
        }

        return array_merge($blocks, $tables, $images, $vars);
    }

    /**
     * @param string $xml
     * @param mixed $search
     * @param mixed $replace
     * @param int $limit
     */
    protected function setValueForScope($xml, $search, $replace, $limit = self::MAXIMUM_REPLACEMENTS_DEFAULT)
    {
        if (is_array($search)) {
            foreach ($search as &$item) {
                $item = static::ensureMacroCompleted($item);
            }
            unset($item);
        } else {
            $search = static::ensureMacroCompleted($search);
        }

        if (is_array($replace)) {
            foreach ($replace as &$item) {
                $item = static::ensureUtf8Encoded($item);
            }
            unset($item);
        } else {
            $replace = static::ensureUtf8Encoded($replace);
            $replace = str_replace('&amp;', '&', $replace);
            $replace = str_replace('&', '&amp;', $replace); // Else the docx generated is corrupted
        }

        // TODO: Enable ???
        // if (Settings::isOutputEscapingEnabled()) {
        //     $xmlEscaper = new Xml();
        //     $replace = $xmlEscaper->escape($replace);
        // }

        return $this->setValueForPart($search, $replace, $xml, $limit);
    }

    /**
     * Clone a table row in a given Xml Scope.
     *
     * @param string $xml
     * @param string $search
     * @param int $numberOfClones
     *
     * @throws \PhpOffice\PhpWord\Exception\Exception
     */
    protected function cloneRowForScope($xml, $search, $numberOfClones)
    {
        $tagPos = strpos($xml, $search);
        if (!$tagPos) {
            // throw new Exception('Can not clone row, template variable not found or variable contains markup.');
            $result = $xml;
        } else {
            $rowStart = $this->findRowStartForScope($xml, $tagPos);
            $rowEnd = $this->findRowEndForScope($xml, $tagPos);
            $xmlRow = $this->getSliceForScope($xml, $rowStart, $rowEnd);

            // Check if there's a cell spanning multiple rows.
            if (preg_match('#<w:vMerge w:val="restart"/>#', $xmlRow)) {
                // $extraRowStart = $rowEnd;
                $extraRowEnd = $rowEnd;
                while (true) {
                    $extraRowStart = $this->findRowStartForScope($xml, $extraRowEnd + 1);
                    $extraRowEnd = $this->findRowEndForScope($xml, $extraRowEnd + 1);

                    // If extraRowEnd is lower then 7, there was no next row found.
                    if ($extraRowEnd < 7) {
                        break;
                    }

                    // If tmpXmlRow doesn't contain continue, this row is no longer part of the spanned row.
                    $tmpXmlRow = $this->getSliceForScope($xml, $extraRowStart, $extraRowEnd);
                    if (!preg_match('#<w:vMerge/>#', $tmpXmlRow) &&
                        !preg_match('#<w:vMerge w:val="continue"\s*/>#', $tmpXmlRow)
                    ) {
                        break;
                    }
                    // This row was a spanned row, update $rowEnd and search for the next row.
                    $rowEnd = $extraRowEnd;
                }
                $xmlRow = $this->getSliceForScope($xml, $rowStart, $rowEnd);
            }

            $result = $this->getSliceForScope($xml, 0, $rowStart);
            $result .= implode('', $this->indexClonedVariables($numberOfClones, $xmlRow));
            $result .= $this->getSliceForScope($xml, $rowEnd);
        }

        return $result;
    }

    /**
     * Get a slice of a string in a given Xml Scope.
     *
     * @param string $xml
     * @param int $startPosition
     * @param int $endPosition
     *
     * @return string
     */
    protected function getSliceForScope($xml, $startPosition, $endPosition = 0)
    {
        if (!$endPosition) {
            $endPosition = strlen($xml);
        }

        return substr($xml, $startPosition, ($endPosition - $startPosition));
    }

    /**
     * Find the end position of the nearest table row after $offset in a given Xml Scope.
     *
     * @param string $xml
     * @param int $offset
     *
     * @return int
     */
    protected function findRowEndForScope($xml, $offset)
    {
        return strpos($xml, '</w:tr>', $offset) + 7;
    }

    /**
     * Find the start position of the nearest table row before $offset in a given Xml Scope.
     *
     * @param string $xml
     * @param int $offset
     *
     * @throws \PhpOffice\PhpWord\Exception\Exception
     *
     * @return int
     */
    protected function findRowStartForScope($xml, $offset)
    {
        $rowStart = strrpos($xml, '<w:tr ', ((strlen($xml) - $offset) * -1));

        if (!$rowStart) {
            $rowStart = strrpos($xml, '<w:tr>', ((strlen($xml) - $offset) * -1));
        }
        if (!$rowStart) {
            throw new Exception('Can not find the start position of the row to clone.');
        }

        return $rowStart;
    }

    /**
     * Replaces variable names in cloned
     * rows/blocks with indexed names
     *
     * /!\ Overide PHPWord to allow replacing images patern variables   EX: ${img#1:30:30}
     *
     * @param int $count
     * @param string $xmlBlock
     *
     * @return array
     */
    protected function indexClonedVariables($count, $xmlBlock)
    {
        $results = array();
        for ($i = 1; $i <= $count; $i++) {
            $results[] = preg_replace('/\$\{(.*?)(:[^}]*)*\}/', '\${${1}#' . $i . '${2}}', $xmlBlock);
        }

        return $results;
    }

    /**
     * @param string $xml
     * @param string $partFileName
     * @param mixed $search
     * @param mixed $replace Path to image, or array("path" => xx, "width" => yy, "height" => zz)
     * @param int $limit
     */
    public function setImageValueForScope($xml, $partFileName, $search, $replace, $limit = self::MAXIMUM_REPLACEMENTS_DEFAULT)
    {
        // prepare $search_replace
        if (!is_array($search)) {
            $search = array($search);
        }

        $replacesList = array();
        if (!is_array($replace) || isset($replace['path'])) {
            $replacesList[] = $replace;
        } else {
            $replacesList = array_values($replace);
        }

        $searchReplace = array();
        foreach ($search as $searchIdx => $searchString) {
            $searchReplace[$searchString] = isset($replacesList[$searchIdx]) ? $replacesList[$searchIdx] : $replacesList[0];
        }

        // define templates
        // result can be verified via "Open XML SDK 2.5 Productivity Tool" (http://www.microsoft.com/en-us/download/details.aspx?id=30425)
        $imgTpl = '<w:pict><v:shape type="#_x0000_t75" style="width:{WIDTH};height:{HEIGHT}"><v:imagedata r:id="{RID}" o:title=""/></v:shape></w:pict>';

        $partVariables = $this->getVariablesForPart($xml);

        foreach ($searchReplace as $searchString => $replaceImage) {
            $varsToReplace = array_filter($partVariables, function ($partVar) use ($searchString) {
                return ($partVar == $searchString) || preg_match('/^' . preg_quote($searchString) . ':/', $partVar);
            });

            foreach ($varsToReplace as $varNameWithArgs) {
                if (empty($replaceImage)) {
                    // remove variable tag
                    $xml = $this->setValueForPart('${' . $varNameWithArgs . '}', "", $xml, $limit);
                } else {
                    $varInlineArgs = $this->getImageArgs($varNameWithArgs);
                    $preparedImageAttrs = $this->prepareImageAttrs($replaceImage, $varInlineArgs);
                    $imgPath = $preparedImageAttrs['src'];

                    // get image index
                    $imgIndex = $this->getNextRelationsIndex($partFileName);
                    $rid = 'rId' . $imgIndex;

                    // replace preparations
                    $this->addImageToRelations($partFileName, $rid, $imgPath, $preparedImageAttrs['mime']);
                    $xmlImage = str_replace(array('{RID}', '{WIDTH}', '{HEIGHT}'), array($rid, $preparedImageAttrs['width'], $preparedImageAttrs['height']), $imgTpl);

                    // replace variable
                    $varNameWithArgsFixed = static::ensureMacroCompleted($varNameWithArgs);
                    $matches = array();
                    if (preg_match('/(<[^<]+>)([^<]*)(' . preg_quote($varNameWithArgsFixed) . ')([^>]*)(<[^>]+>)/Uu', $xml, $matches)) {
                        $wholeTag = $matches[0];
                        array_shift($matches);
                        list($openTag, $prefix,, $postfix, $closeTag) = $matches;
                        $replaceXml = $openTag . $prefix . $closeTag . $xmlImage . $openTag . $postfix . $closeTag;
                        // replace on each iteration, because in one tag we can have 2+ inline variables => before proceed next variable we need to change $xml
                        $xml = $this->setValueForPart($wholeTag, $replaceXml, $xml, $limit);
                    }
                }
            }
        }

        return $xml;
    }

    // No change from PHPWord exept being protected rather than private...
    protected function getImageArgs($varNameWithArgs)
    {
        $varElements = explode(':', $varNameWithArgs);
        array_shift($varElements); // first element is name of variable => remove it

        $varInlineArgs = array();
        // size format documentation: https://msdn.microsoft.com/en-us/library/documentformat.openxml.vml.shape%28v=office.14%29.aspx?f=255&MSPPError=-2147217396
        foreach ($varElements as $argIdx => $varArg) {
            if (strpos($varArg, '=')) { // arg=value
                list($argName, $argValue) = explode('=', $varArg, 2);
                $argName = strtolower($argName);
                if ($argName == 'size') {
                    list($varInlineArgs['width'], $varInlineArgs['height']) = explode('x', $argValue, 2);
                } else {
                    $varInlineArgs[strtolower($argName)] = $argValue;
                }
            } elseif (preg_match('/^([0-9]*[a-z%]{0,2}|auto)x([0-9]*[a-z%]{0,2}|auto)$/i', $varArg)) { // 60x40
                list($varInlineArgs['width'], $varInlineArgs['height']) = explode('x', $varArg, 2);
            } else { // :60:40:f
                switch ($argIdx) {
                    case 0:
                        $varInlineArgs['width'] = $varArg;
                        break;
                    case 1:
                        $varInlineArgs['height'] = $varArg;
                        break;
                    case 2:
                        $varInlineArgs['ratio'] = $varArg;
                        break;
                }
            }
        }

        return $varInlineArgs;
    }

    // No change from PHPWord exept being protected rather than private...
    protected function prepareImageAttrs($replaceImage, $varInlineArgs)
    {
        // get image path and size
        $width = null;
        $height = null;
        $ratio = null;
        if (is_array($replaceImage) && isset($replaceImage['path'])) {
            $imgPath = $replaceImage['path'];
            if (isset($replaceImage['width'])) {
                $width = $replaceImage['width'];
            }
            if (isset($replaceImage['height'])) {
                $height = $replaceImage['height'];
            }
            if (isset($replaceImage['ratio'])) {
                $ratio = $replaceImage['ratio'];
            }
        } else {
            $imgPath = $replaceImage;
        }

        $width = $this->chooseImageDimension($width, isset($varInlineArgs['width']) ? $varInlineArgs['width'] : null, 115);
        $height = $this->chooseImageDimension($height, isset($varInlineArgs['height']) ? $varInlineArgs['height'] : null, 70);

        $imageData = @getimagesize($imgPath);
        if (!is_array($imageData)) {
            throw new Exception(sprintf('Invalid image: %s', $imgPath));
        }
        list($actualWidth, $actualHeight, $imageType) = $imageData;

        // fix aspect ratio (by default)
        if (is_null($ratio) && isset($varInlineArgs['ratio'])) {
            $ratio = $varInlineArgs['ratio'];
        }
        if (is_null($ratio) || !in_array(strtolower($ratio), array('', '-', 'f', 'false'))) {
            $this->fixImageWidthHeightRatio($width, $height, $actualWidth, $actualHeight);
        }

        $imageAttrs = array(
            'src'    => $imgPath,
            'mime'   => image_type_to_mime_type($imageType),
            'width'  => $width,
            'height' => $height,
        );

        return $imageAttrs;
    }

    // No change from PHPWord exept being protected rather than private...
    protected function chooseImageDimension($baseValue, $inlineValue, $defaultValue)
    {
        $value = $baseValue;
        if (is_null($value) && isset($inlineValue)) {
            $value = $inlineValue;
        }
        if (!preg_match('/^([0-9]*(cm|mm|in|pt|pc|px|%|em|ex|)|auto)$/i', $value)) {
            $value = null;
        }
        if (is_null($value)) {
            $value = $defaultValue;
        }
        if (is_numeric($value)) {
            $value .= 'px';
        }

        return $value;
    }

    // No change from PHPWord exept being protected rather than private...
    protected function fixImageWidthHeightRatio(&$width, &$height, $actualWidth, $actualHeight)
    {
        $imageRatio = $actualWidth / $actualHeight;

        if (($width === '') && ($height === '')) { // defined size are empty
            $width = $actualWidth . 'px';
            $height = $actualHeight . 'px';
        } elseif ($width === '') { // defined width is empty
            $heightFloat = (float) $height;
            $widthFloat = $heightFloat * $imageRatio;
            $matches = array();
            preg_match("/\d([a-z%]+)$/", $height, $matches);
            $width = $widthFloat . $matches[1];
        } elseif ($height === '') { // defined height is empty
            $widthFloat = (float) $width;
            $heightFloat = $widthFloat / $imageRatio;
            $matches = array();
            preg_match("/\d([a-z%]+)$/", $width, $matches);
            $height = $heightFloat . $matches[1];
        } else { // we have defined size, but we need also check it aspect ratio
            $widthMatches = array();
            preg_match("/\d([a-z%]+)$/", $width, $widthMatches);
            $heightMatches = array();
            preg_match("/\d([a-z%]+)$/", $height, $heightMatches);
            // try to fix only if dimensions are same
            if ($widthMatches[1] == $heightMatches[1]) {
                $dimention = $widthMatches[1];
                $widthFloat = (float) $width;
                $heightFloat = (float) $height;
                $definedRatio = $widthFloat / $heightFloat;

                if ($imageRatio > $definedRatio) { // image wider than defined box
                    $height = ($widthFloat / $imageRatio) . $dimention;
                } elseif ($imageRatio < $definedRatio) { // image higher than defined box
                    $width = ($heightFloat * $imageRatio) . $dimention;
                }
            }
        }
    }

    // No change from PHPWord exept being protected rather than private...
    protected function addImageToRelations($partFileName, $rid, $imgPath, $imageMimeType)
    {
        // define templates
        $typeTpl = '<Override PartName="/word/media/{IMG}" ContentType="image/{EXT}"/>';
        $relationTpl = '<Relationship Id="{RID}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/{IMG}"/>';
        $newRelationsTpl = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' . "\n" . '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
        $newRelationsTypeTpl = '<Override PartName="/{RELS}" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>';
        $extTransform = array(
            'image/jpeg' => 'jpeg',
            'image/png'  => 'png',
            'image/bmp'  => 'bmp',
            'image/gif'  => 'gif',
        );

        // get image embed name
        if (isset($this->tempDocumentNewImages[$imgPath])) {
            $imgName = $this->tempDocumentNewImages[$imgPath];
        } else {
            // transform extension
            if (isset($extTransform[$imageMimeType])) {
                $imgExt = $extTransform[$imageMimeType];
            } else {
                throw new Exception("Unsupported image type $imageMimeType");
            }

            // add image to document
            $imgName = 'image_' . $rid . '_' . pathinfo($partFileName, PATHINFO_FILENAME) . '.' . $imgExt;
            $this->zipClass->pclzipAddFile($imgPath, 'word/media/' . $imgName);
            $this->tempDocumentNewImages[$imgPath] = $imgName;

            // setup type for image
            $xmlImageType = str_replace(array('{IMG}', '{EXT}'), array($imgName, $imgExt), $typeTpl);
            $this->tempDocumentContentTypes = str_replace('</Types>', $xmlImageType, $this->tempDocumentContentTypes) . '</Types>';
        }

        $xmlImageRelation = str_replace(array('{RID}', '{IMG}'), array($rid, $imgName), $relationTpl);

        if (!isset($this->tempDocumentRelations[$partFileName])) {
            // create new relations file
            $this->tempDocumentRelations[$partFileName] = $newRelationsTpl;
            // and add it to content types
            $xmlRelationsType = str_replace('{RELS}', $this->getRelationsName($partFileName), $newRelationsTypeTpl);
            $this->tempDocumentContentTypes = str_replace('</Types>', $xmlRelationsType, $this->tempDocumentContentTypes) . '</Types>';
        }

        // add image to relations
        $this->tempDocumentRelations[$partFileName] = str_replace('</Relationships>', $xmlImageRelation, $this->tempDocumentRelations[$partFileName]) . '</Relationships>';
    }
}
