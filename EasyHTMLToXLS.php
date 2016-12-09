<?php
/*!
 * EasyHTMLToXLS HTML parser for PHPExcel v0.8
 * https://github.com/andrehtissot/easy-html-to-xls
 *
 * Requires PHPExcel
 * http://phpexcel.codeplex.com/
 *
 * Copyright André Augusto Tissot
 * Released under the MIT license
 *
 * Date: 2016-12-09
 */
class EasyHTMLToXLS {
    protected $debugMode = true;
    protected $currentSheetIndex = 0;
    public $sheet = null;
    protected $objPHPExcel = null;
    protected $hasWrittenFirstLineBySheetIndex = array(0 => false);
    protected $styles = array('th' => array(), 'thead' => array(), 'tbody' => array(),
        'tfoot' => array(), 'td' => array(), 'h1' => array(), 'h2' => array(),
        'h3' => array(), 'h4' => array(), 'h5' => array(), 'h6' => array(),
        'img' => array(), 'strong' => array(), 'table' => array());
    protected $stringWidthsMatrixBySheetIndex = array();
    protected $columnWidthBySheetIndex = array();
    protected $oneCellRowsBySheetIndex = array();
    protected $rowHeightBySheetIndex = array();

    public static function stream($html, $file_name, array $options = array()){
        $easyxls = new self();
        if(!isset($options['title'])) { $options['title'] = str_replace("\n", ' ', $file_name); }
        $objWriter = $easyxls->generate($html, $options['title'], $options);
        header('Content-Type: application/vnd.ms-excel');
        header("Content-Disposition: attachment;filename=\"$file_name.xls\"");
        header('Cache-Control: max-age=0');
        $objWriter->save('php://output');
        exit;
    }
    public static function write($html, $title, $output_file_path, array $options = array()){
        $easyxls = new self();
        $objWriter = $easyxls->generate($html, $title, $options);
        $objWriter->save($output_file_path);
    }

    public function get_letters($column_number){
        $alphabet = array('', 'A','B','C','D','E','F','G','H','I','J','K',
            'L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z');
        if($column_number === 0){
            debug_print_backtrace();
            die;
        }
        $column_number -= 1;
        return $alphabet[floor($column_number/26)].$alphabet[($column_number % 26)+1];
    }

    protected function get_column_number($letters){
        $alphabet = array('A','B','C','D','E','F','G','H','I','J','K',
            'L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z');
        $result = 0;
        $letters_array = array_reverse(str_split($letters));
        $number = 1+array_search($letters_array[0], $alphabet);
        $number += isset($letters_array[1]) ? (26*(1+array_search($letters_array[1], $alphabet))) : 0;
        return $number;
    }

    protected function next_row_number($sheetIndex = null){
        if(is_null($sheetIndex)){
            $sheet = $this->sheet;
            $sheetIndex = $this->currentSheetIndex;
        } else $sheet = $this->objPHPExcel->getSheet($sheetIndex);
        if($this->hasWrittenFirstLineBySheetIndex[$sheetIndex] || $sheet->getHighestRow() !== 1)
            return $sheet->getHighestRow()+1;
        return 1;
    }

    protected function setCellValue($sheet, $column_number, $row_number, $value){
        if(trim($value) === '' && !$this->hasWrittenFirstLineBySheetIndex[$this->currentSheetIndex]) { return; }
        $this->hasWrittenFirstLineBySheetIndex[$this->currentSheetIndex] = true;
        if(!isset($this->stringWidthsMatrixBySheetIndex[$this->currentSheetIndex]))
            $this->stringWidthsMatrixBySheetIndex[$this->currentSheetIndex] = array();
        if(!isset($this->stringWidthsMatrixBySheetIndex[$this->currentSheetIndex][$column_number]))
            $this->stringWidthsMatrixBySheetIndex[$this->currentSheetIndex][$column_number] = array(
                $row_number => mb_strwidth(trim($value)));
        else $this->stringWidthsMatrixBySheetIndex[$this->currentSheetIndex][$column_number][$row_number]
            = max((int) @$this->stringWidthsMatrixBySheetIndex[$this->currentSheetIndex][$column_number][$row_number],
                mb_strwidth(trim($value)));
        $sheet->setCellValue(self::get_letters($column_number) . $row_number, trim($value));
    }

    protected function getCellStyle($sheet, $column_number, $row_number){
        return $sheet->getStyle(self::get_letters($column_number).$row_number);
    }
    static protected function getProtectedValue($obj, $name) {
      $array = (array) $obj;
      $prefix = chr(0).'*'.chr(0);
      return $array[$prefix.$name];
    }

    protected function getCellValue($sheet, $column_number, $row_number){
        return $sheet->getCell(self::get_letters($column_number).$row_number)->getValue();
    }

    protected function setWidthFromNodeFromChildren($node, &$cssStyleAttrs){
        if(!empty($cssStyleAttrs['width'])) { return; }
        if(!$node->hasChildNodes()) { return; }
        $minWidth = @$cssStyleAttrs['min-width'];
        $maxWidth = @$cssStyleAttrs['max-width'];
        if($minWidth !== null && $maxWidth !== null) { return; }
        $childrenMinWidth = null;
        $childrenMaxWidth = null;
        $children = $node->childNodes;
        foreach ($children as $child) {
            if(is_a($child, 'DOMText')){
                if(trim($child->wholeText) !== '') { return; }
                continue;
            }
            $childCssStyleAttrs = self::parseCSSAttributes($child->getAttribute('style'));
            $this->setWidthFromNodeFromChildren($child, $childCssStyleAttrs);
            if(!empty($childCssStyleAttrs['width'])) {
                $childrenMinWidth+=$this->convertSizeUnit($childCssStyleAttrs['width'], $child->tagName);
                $childrenMaxWidth+=$this->convertSizeUnit($childCssStyleAttrs['width'], $child->tagName);
                continue;
            }
            if(!empty($childCssStyleAttrs['min-width']))
                $childrenMinWidth+=$this->convertSizeUnit($childCssStyleAttrs['min-width'], $child->tagName);
            if(!empty($childCssStyleAttrs['max-width']))
                $childrenMaxWidth+=$this->convertSizeUnit($childCssStyleAttrs['max-width'], $child->tagName);
            elseif(empty($childCssStyleAttrs['min-width']))
                return;
        }
        if($childrenMinWidth === $childrenMaxWidth){
            $cssStyleAttrs['width'] = $childrenMaxWidth;
            return;
        }
        if($childrenMaxWidth !== null)
            $cssStyleAttrs['max-width'] = $childrenMaxWidth;
        if($childrenMinWidth !== null)
            $cssStyleAttrs['min-width'] = $childrenMinWidth;
    }

    protected function setAutomaticStyleFromNode($dom_element, $initial_column_number, $initial_row_number,
        $final_column_number = null, $final_row_number = null){
        $cssStyleAttrs = self::parseCSSAttributes($dom_element->getAttribute('style'));
        $this->setWidthFromNodeFromChildren($dom_element, $cssStyleAttrs);
        $this->setAutomaticStyle($dom_element->tagName, $dom_element->textContent, $initial_column_number,
            $initial_row_number, $final_column_number, $final_row_number, $cssStyleAttrs);
    }

    protected function setAutomaticStyle($tag_name, $value, $initial_column_number, $initial_row_number,
        $final_column_number = null, $final_row_number = null, $cssStyleAttrs = null){
        if(!is_numeric($final_column_number) || $final_column_number < $initial_column_number)
            $final_column_number = $initial_column_number;
        if(!is_numeric($final_row_number) || $final_row_number < $initial_row_number)
            $final_row_number = $initial_row_number;
        if($tag_name === 'th' || $tag_name === 'td'){
            $value = trim($value);
            if(isset($cssStyleAttrs['cell-format']) && $cssStyleAttrs['cell-format'] === 'text'){
                $style = $this->styles[$tag_name];
            } elseif(preg_match('/^R\$ ?-?\d?\d?\d(\.\d{3})*[,\.]\d\d$/', $value))
                $style = $this->styles["{$tag_name}_monetary"];
            elseif(preg_match('/^-?\d?\d?\d(\.\d{3})*[,\.]\d\d ?%$/', $value))
                $style = $this->styles["{$tag_name}_relative"];
            elseif(preg_match('/^-?\d?\d?\d(\.\d{3})*[,\.]\d\d$/', $value))
                $style = $this->styles["{$tag_name}_decimal"];
            else $style = $this->styles[$tag_name];
            if(!empty($cssStyleAttrs))
                $style = $this->applyCSSToDefaultStyle($cssStyleAttrs, $style, $tag_name);
            if(empty($style)) { return; }
            for ($column_number=$initial_column_number; $column_number <= $final_column_number; $column_number++){
                for ($row_number=$initial_row_number; $row_number <= $final_row_number; $row_number++)
                    $this->sheet->getStyle(self::get_letters($column_number).$row_number)->applyFromArray($style);
            }
            $this->setWitdhFromCssStyle($cssStyleAttrs, $initial_column_number, $tag_name);
            return;
        }
        if(empty($cssStyleAttrs) && empty($this->styles[$tag_name])) { return; }
        $styles = empty($this->styles[$tag_name]) ? array() : $this->styles[$tag_name];
        if(!empty($cssStyleAttrs))
            $styles = $this->applyCSSToDefaultStyle($cssStyleAttrs, $styles, $tag_name);
        if(is_array($styles)) {
            for ($column_number=$initial_column_number; $column_number <= $final_column_number; $column_number++)
                for ($row_number=$initial_row_number; $row_number <= $final_row_number; $row_number++)
                    $this->sheet->getStyle(self::get_letters($column_number).$row_number)->applyFromArray($styles);
            return;
        }
        if(is_a($styles, 'Closure')){
            $styles($this, $value, $initial_column_number, $initial_row_number, $final_column_number,
                $final_row_number);
            return;
        }
        if($this->debugMode){
            p($this->styles);
            die(p($tag_name));
        }
    }

    protected function convertSizeUnitFromPX($value){
        return floatval($value)/8;
    }

    protected function convertSizeUnit($valueAsString, $tagName = null){
        if($valueAsString === null) { return null; }
        if(is_numeric($valueAsString)) { return $valueAsString; }
        if(strpos($valueAsString, 'px') !== false)
            return $this->convertSizeUnitFromPX(substr($valueAsString,0,-2));
        if($this->debugMode){ echo 'Conversion required: '; p($valueAsString);die(); }
    }

    protected function setWitdhFromCssStyle($styleAttrs, $columnNumber, $tagName){
        if(empty($styleAttrs['width']) && empty($styleAttrs['min-width']) && empty($styleAttrs['max-width'])){ return; }
        if(!isset($this->columnWidthBySheetIndex[$this->currentSheetIndex]))
            $this->columnWidthBySheetIndex[$this->currentSheetIndex] = array();
        if($styleAttrs['width']){
            $this->columnWidthBySheetIndex[$this->currentSheetIndex][$columnNumber]
                = $this->convertSizeUnit($styleAttrs['width'], $tagName);
            return;
        }
        $this->columnWidthBySheetIndex[$this->currentSheetIndex][$columnNumber]
            = array($this->convertSizeUnit(@$styleAttrs['min-width'], $tagName),
                $this->convertSizeUnit(@$styleAttrs['max-width'], $tagName));
    }

    protected function generate($html, $title, array $options = array()){
        $this->objPHPExcel = new PHPExcel();
        $this->objPHPExcel->setActiveSheetIndex(0);
        $this->sheet = $this->objPHPExcel->getActiveSheet();
        $this->loadLayout('begin', $options);
        $this->initStyle($options);
        self::exitIfInvalidHTML($html);
        $this->loadAndFixHtml($html);
        // die;
        $this->defineColumnWidths();
        $this->defineColumnHeights();
        // die;
        $this->loadLayout('end', $options);
        $this->objPHPExcel->setActiveSheetIndex(0);
        return PHPExcel_IOFactory::createWriter($this->objPHPExcel, 'Excel5');
    }

    protected function loadLayout($layoutName, &$options){
        $layout = isset($options["layout_$layoutName"]) ? $options["layout_$layoutName"]
            : self::getConfig("defaults.layout_{$layoutName}_path");
        $layout = (strpos($layout, '/') === 0) ? __DIR__."/../../../..$layout" : $layout;
        if(file_exists($layout)) { include $layout; }
    }

    protected function initStyle(&$options){
        if(isset($options['styles']))
            $this->styles = array_replace_recursive($this->styles, $options['styles']);
        unset($options['styles']);
        foreach(array('th','td') as $tagName)
            foreach(array('monetary' => PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_BR_SIMPLE,
                'relative' => PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE_00,
                'decimal' => PHPExcel_Style_NumberFormat::FORMAT_NUMBER_00) as $variant => $excelStyle){
                $this->styles[$tagName.'_'.$variant] = $this->styles[$tagName];
                $this->styles[$tagName.'_'.$variant]['numberformat'] = array('code' => $excelStyle);
            }
    }

    protected function loadAndFixHtml($html){
        libxml_use_internal_errors(true);
        $dom_document = new DOMDocument();
        if(strpos($html, '<html>') === false) { $html = "<html>$html"; }
        if(strpos($html, '<?xml encoding="UTF-8">') === false) { $html = '<?xml encoding="UTF-8">'.$html; }
        if(strpos($html, '</html>') === false) { $html = "$html</html>"; }
        @$dom_document->loadHTML($html);
        foreach ($dom_document->childNodes as $item)
            if ($item->nodeType == XML_PI_NODE)
                $dom_document->removeChild($item);
        $dom_document->encoding = 'UTF-8';
        $this->generate_node($dom_document);
        $dom_document = null;
        libxml_use_internal_errors(false);
        libxml_use_internal_errors(true);
        flush();
    }

    protected function defineColumnHeights(){
        foreach ($this->objPHPExcel->getAllSheets() as $sheetIndex => $sheet)
            if(!empty($this->rowHeightBySheetIndex[$sheetIndex]))
                foreach ($this->rowHeightBySheetIndex[$sheetIndex] as $rowNumber => $height){
                    if($height === -1){
                        // $maxHeight = null;
                        // for($columnNumber = self::get_column_number($sheet->getHighestColumn()); $columnNumber > 0;
                        //     $columnNumber--){
                        //     $maxHeight = max($maxHeight, $this->getIdealHeightForCell($sheetIndex, $columnNumber,
                        //         $rowNumber));
                        //         // $this->calculateHightFromText(
                        //         // $this->getCellValue($sheet, $columnNumber, $rowNumber),
                        //         // $sheet->getStyle(self::get_letters($columnNumber).$rowNumber)->getFont()->getSize(),
                        //         // $sheet->getColumnDimension(self::get_letters($columnNumber))->getWidth()));
                        // }
                        // if($maxHeight !== null)
                        //     $sheet->getRowDimension($rowNumber)->setRowHeight($maxHeight);
                    } else $sheet->getRowDimension($rowNumber)->setRowHeight($height);
                }
    }

    protected function calculateHightFromText($text, $fontSize, $columnWidth){
        $encode = ini_get('default_charset');
        $linesCounter = null;
        $lines = explode("\n", $text);
        foreach ($lines as $line)
            $linesCounter += ceil(iconv_strlen($line, $encode) / $columnWidth);
        if($linesCounter === null) { $linesCounter = 1; }
        // list($left,, $right) = imageftbbox( 12, 0, arial.ttf, "Hello World");
        return $this->convertFontSizeToHeight($fontSize * $linesCounter);
    }

    protected function defineColumnWidths(){
        foreach ($this->objPHPExcel->getAllSheets() as $sheetIndex => $sheet) {
            $lastColumnLettersIndex = self::get_column_number($sheet->getHighestColumn());
            if(!empty($this->oneCellRowsBySheetIndex[$sheetIndex]))
                foreach ($this->oneCellRowsBySheetIndex[$sheetIndex] as $oneCellRow)
                    $this->mergeCells($sheetIndex, 1, $oneCellRow, $lastColumnLettersIndex, $oneCellRow);
            for($i = self::get_column_number($sheet->getHighestColumn()); $i > 0; $i--)
                if(empty($this->columnWidthBySheetIndex[$sheetIndex][$i])
                    || is_array($this->columnWidthBySheetIndex[$sheetIndex][$i]))
                $sheet->getColumnDimension(self::get_letters($i))->setAutoSize(true);
            $sheet->calculateColumnWidths();
            for($i = self::get_column_number($sheet->getHighestColumn()); $i > 0; $i--){
                if(@$this->columnWidthBySheetIndex[$sheetIndex][$i]
                    && !is_array($this->columnWidthBySheetIndex[$sheetIndex][$i])){
                    $width = $this->columnWidthBySheetIndex[$sheetIndex][$i];
                } else {
                    $width = $sheet->getColumnDimension(self::get_letters($i))->getWidth();
                    if(!isset($this->stringWidthsMatrixBySheetIndex[$sheetIndex]))
                        $this->stringWidthsMatrixBySheetIndex[$sheetIndex] = array();
                    $width = max(max(empty($this->stringWidthsMatrixBySheetIndex[$sheetIndex][$i]) ? array(0)
                        : $this->stringWidthsMatrixBySheetIndex[$sheetIndex][$i]), $width);
                    if(@$this->columnWidthBySheetIndex[$sheetIndex][$i]) {
                        if($this->columnWidthBySheetIndex[$sheetIndex][$i][0] !== null &&
                            $this->columnWidthBySheetIndex[$sheetIndex][$i][0] > $width)
                            $width = $this->columnWidthBySheetIndex[$sheetIndex][$i][0];
                        elseif($this->columnWidthBySheetIndex[$sheetIndex][$i][1] !== null &&
                            $this->columnWidthBySheetIndex[$sheetIndex][$i][1] < $width)
                            $width = $this->columnWidthBySheetIndex[$sheetIndex][$i][1];
                    }
                }
                if(@$this->columnWidthBySheetIndex[$sheetIndex][$i])
                    for($rowNumber = $this->next_row_number()-1; $rowNumber > 0; $rowNumber--)
                        $this->setWrapText($sheetIndex, $i, $rowNumber);
                $sheet->getColumnDimension(self::get_letters($i))->setAutoSize(false)->setWidth($width);
            }
        }
    }

    protected function setWrapText($sheetIndex, $columnNumber, $rowNumber){
        $this->objPHPExcel->getSheet($sheetIndex)->getStyle(self::get_letters($columnNumber).$rowNumber)
            ->getAlignment()->setWrapText(true);
        if(empty($this->rowHeightBySheetIndex[$sheetIndex]))
            $this->rowHeightBySheetIndex[$sheetIndex] = array($rowNumber => -1);
        elseif(empty($this->rowHeightBySheetIndex[$sheetIndex][$rowNumber]))
            $this->rowHeightBySheetIndex[$sheetIndex][$rowNumber] = -1;
    }

    protected $non_defining_parent_nodes = array('body','font','label','a','div','span','strong','b','p','html');
    protected function has_any_defining_child_node($node){
        if($node === null || !$node->hasChildNodes()) { return false; }
        foreach ($node->childNodes as $childNode)
            if(is_a($childNode, 'DOMElement'))
                if(!in_array($childNode->tagName, $this->non_defining_parent_nodes)
                    || $this->has_any_defining_child_node($childNode))
                    return true;
        return false;
    }

    protected function generate_node($dom_element_node){
        if(is_a($dom_element_node, 'DOMText')) {
            if(trim($dom_element_node->wholeText) === '') { return; }
            $generateNodeHandledGeneration = 0;
            do {
                do {
                    $parentNode = $dom_element_node->parentNode;
                    if($parentNode->tagName === 'html' || $parentNode->tagName === 'body'){
                        $this->generateNodeCloserToBody($dom_element_node);
                        return array('handledGenerations' => $generateNodeHandledGeneration);
                    }
                    if($this->has_any_defining_child_node($parentNode)) { break; }
                    $dom_element_node = $parentNode;
                    $generateNodeHandledGeneration++;
                } while(is_a($parentNode, 'DOMElement'));
                $siblings = array();
                if($parentNode && $parentNode->hasChildNodes())
                    foreach ($parentNode->childNodes as $childNode)
                        $siblings[] = $childNode;
                $currentNodeIndex = array_search($dom_element_node, $siblings);
                $isFirstNode = true;
                if($currentNodeIndex !== 0){
                    for ($i=$currentNodeIndex-1; $i >= 0; $i--) {
                        if(is_a($siblings[$i], 'DOMText')){
                            if(trim($siblings[$i]->textContent) === '') {
                                if($i === 0) break 2;
                                else continue;
                            } else return array('handledGenerations' => $generateNodeHandledGeneration);
                        }
                        if(is_a($siblings[$i], 'DOMElement')){
                            if(!in_array($siblings[$i]->tagName, $this->non_defining_parent_nodes)
                                || $this->has_any_defining_child_node($siblings[$i])){
                                $isFirstNode = $i === 0;
                                break 2;
                            }
                            return array('handledGenerations' => $generateNodeHandledGeneration);
                        }
                    }
                }
                $siblingsCount = count($siblings);
                $text = trim($dom_element_node->textContent);
                for ($i=$currentNodeIndex+1; $i < $siblingsCount; $i++) {
                    if(is_a($siblings[$i], 'DOMElement')
                        && !in_array(@$siblings[$i]->tagName, $this->non_defining_parent_nodes))
                        break;
                    $text = trim("$text ".trim($siblings[$i]->textContent));
                }
                $dom_element_node = $parentNode;
            } while($i === $siblingsCount && $isFirstNode);
            self::generate_one_column_text($text, $parentNode->tagName);
            return array('handledGenerations' => $generateNodeHandledGeneration);
        }
        if(is_a($dom_element_node, 'DOMElement')
            && !in_array(@$dom_element_node->tagName, $this->non_defining_parent_nodes)){
            switch (@$dom_element_node->tagName) {
                case 'style': return;
                case 'head': return;
                case 'h1':
                case 'h2':
                case 'h3':
                case 'h4':
                case 'h5':
                case 'h6': self::generate_one_column_text_from_node($dom_element_node); return;
                case 'table': self::generate_table($dom_element_node); return;
                case 'br': self::generate_br(); return;
                case 'tab': self::handleTab($dom_element_node); break;
                default:
                    if($this->debugMode){
                        echo 'Node not found on generate_node: ';
                        die(p($dom_element_node));
                    }
            }
        }
        if($this->hasStyle($dom_element_node))
            $initial_row_number = $this->next_row_number();
        $handledGenerations = 0;
        if($dom_element_node->hasChildNodes()){
            foreach ($dom_element_node->childNodes as $idx => $node){
                $return = self::generate_node($node);
                if(!empty($return['handledGenerations'])){
                    $handledGenerations = $return['handledGenerations'] - 1;
                    break;
                }
            }
        }
        if($this->hasStyle($dom_element_node)){
            $final_row_number = $this->next_row_number()-1;
            $this->setAutomaticStyleFromNode($dom_element_node, 1, $initial_row_number, 1, $final_row_number);
        }
        if($handledGenerations)
            return array('handledGenerations' => $handledGenerations);
    }

    protected function generateNodeCloserToBody($domElementNode){
        self::generate_one_column_text_from_node(self::removeInternalTags($domElementNode,
            array('style','head','table','tab','h1','h2','h3','h4','h5','h6')));
    }

    protected function generate_table($dom_element_table){
        $initial_row_number = $this->next_row_number();
        $row_number = $initial_row_number - 1;
        $first_column_number = $column_number = 0;
        $occupied_matrix = array();
        if($dom_element_table->hasAttribute('style')){
            $tableStyle = self::parseCSSAttributes($dom_element_table->getAttribute('style'));
            if(isset($tableStyle['initial-column']))
                $column_number = $tableStyle['initial-column'];
            $first_column_number = $column_number = (int) $column_number;
        }
        $this->generate_table_children($dom_element_table, $column_number, $row_number, $occupied_matrix);
    }

    protected function generate_table_children($dom_element_table, &$first_column_number, &$row_number, &$occupied_matrix){
        if(is_a($dom_element_table, 'DOMText') && trim($dom_element_table->wholeText, " \r\n\t") === '') { continue; }
        if(!$dom_element_table->hasChildNodes()) { return; }
        $column_number = $first_column_number;
        $childNodes = ($dom_element_table->tagName === 'th' || $dom_element_table->tagName === 'td')
            ? array($dom_element_table) : $dom_element_table->childNodes;
        foreach ($childNodes as $node) {
            if(is_a($node, 'DOMText') && trim($node->wholeText) === '') { continue; }
            if($node->tagName === 'tr'){
                $row_number++;
                $this->generate_table_children($node, $column_number, $row_number, $occupied_matrix);
            } elseif($node->tagName === 'thead' || $node->tagName === 'tbody' || $node->tagName === 'tfoot'){
                $has_tr_children = false;
                foreach ($node->childNodes as $childNode) {
                    if(@$childNode->tagName === 'tr'){
                        $row_number++;
                        $this->generate_table_children($childNode, $column_number, $row_number, $occupied_matrix);
                        $has_tr_children = true;
                    }
                }
                if(!$has_tr_children){
                    $row_number++;
                    $this->generate_table_children($node, $column_number, $row_number, $occupied_matrix);
                }
                if(!empty($occupied_matrix[$row_number]))
                    $this->setAutomaticStyle($node->tagName, '', $first_column_number+1, $row_number,
                        max($occupied_matrix[$row_number]), $row_number);
            } elseif($node->tagName === 'th' || $node->tagName === 'td'){
                $column_number++;
                $this->generate_table_tx($node, $row_number, $column_number, $occupied_matrix);
            } elseif($this->debugMode){
                echo 'generate_table_children: ';
                var_dump($node->wholeText);
                die(p($node));
            }
        }
    }

    protected function generate_table_tx_cell_content($tx_dom_element, $row_number, $column_number){
        foreach ($tx_dom_element->childNodes as $node) {
            if(isset($node->tagName) && $node->tagName == 'img') {
                foreach (array('alt','title','src','width','height') as $var)
                    $$var = $node->hasAttribute($var) ? $node->getAttribute($var) : '';
                $src = strpos($src, self::getConfig('application_route')) === 0
                    ? '.'.substr($src, strlen(self::getConfig('application_route'))) : $src;
                $this->setAutomaticStyleFromNode($node, $column_number, $row_number);
                if(empty($src)) break;
                $objDrawing = new PHPExcel_Worksheet_Drawing();
                $objDrawing->setName($title);
                $objDrawing->setDescription($alt);
                $objDrawing->setPath($src);
                $objDrawing->setCoordinates(self::get_letters($column_number) . $row_number);
                if(!empty($height) && !empty($width))
                    $objDrawing->setResizeProportional(false);
                if(!empty($height)){
                    $objDrawing->setHeight((int) $height);
                    $this->sheet->getRowDimension($row_number)->setRowHeight(
                        max($this->sheet->getRowDimension($row_number)->getRowHeight(), (int) 1233));
                }
                if(!empty($width)){
                    $objDrawing->setWidth((int) $width);
                    $columnDimension = $this->sheet->getColumnDimension(self::get_letters($column_number));
                    $columnDimension->setAutoSize(false);
                    $columnDimension->setWidth(max($columnDimension->getWidth(),(int) $width));
                }
                $objDrawing->setWorksheet($this->sheet);
            }
        }
        $textContent = preg_replace('/  +/', ' ', trim($tx_dom_element->textContent));
        $this->setAutomaticStyle($tx_dom_element->tagName, $textContent, $column_number,
            $row_number);
        if(preg_match('/^R\$ ?-?\d?\d?\d(\.\d{3})*,\d\d$/', $textContent))
            return self::real_to_value(str_replace(array('R$',' '), '', $textContent));
        elseif(preg_match('/^-?\d?\d?\d(\.\d{3})*,\d\d ?%$/', $textContent))
            return self::real_to_value(str_replace(array('%',' '), '', $textContent))/100;
        elseif(preg_match('/^-?\d?\d?\d(\.\d{3})*,\d\d$/', $textContent))
            return self::real_to_value($textContent);
        return $textContent;
    }

    protected function generate_table_tx($tx_dom_element, &$row_number, &$column_number, &$occupied_matrix){
        while(@in_array($column_number, $occupied_matrix[$row_number]))
            $column_number++;
        $this->setCellValue($this->sheet, $column_number, $row_number,
            $this->generate_table_tx_cell_content($tx_dom_element, $row_number, $column_number));
        $this->setAutomaticStyleFromNode($tx_dom_element, $column_number, $row_number);
        $rowspan = $tx_dom_element->hasAttribute('rowspan') ? $tx_dom_element->getAttribute('rowspan') : 1;
        $colspan = $tx_dom_element->hasAttribute('colspan') ? $tx_dom_element->getAttribute('colspan') : 1;
        for ($add_to_row=0; $add_to_row < $rowspan; $add_to_row++){
            for ($add_to_column=0; $add_to_column < $colspan; $add_to_column++){
                if($add_to_row+$add_to_column !== 0)
                    $this->setAutomaticStyleFromNode($tx_dom_element, $column_number+$add_to_column, $row_number+$add_to_row);
                if(!isset($occupied_matrix[$row_number+$add_to_row]))
                    $occupied_matrix[$row_number+$add_to_row] = array();
                $occupied_matrix[$row_number+$add_to_row][] = $column_number+$add_to_column;
            }
        }
        if($rowspan + $colspan > 2)
            $this->mergeCells($this->currentSheetIndex, $column_number, $row_number, $column_number+$colspan-1, $row_number+$rowspan-1);
        $column_number+=($colspan-1);
    }

    protected function real_to_value($value) {
        $value = trim($value);
        $value = strpos($value, ',') !== false ? str_replace(',', '.', str_replace('.', '', $value)) : $value;
        return empty($value) ? 0 : $value;
    }

    protected function generate_br(){
        $row_number = $this->next_row_number();
        $this->setCellValue($this->sheet, 1, $row_number, '');
    }

    protected function getHightFromTextAndTagName($text, $tagName){
        $this->getFontHeightFromTagName($tagName) * (1+substr_count($text, "\n"));
    }

    protected function generate_one_column_text($text, $tagName){
        $row_number = $this->next_row_number();
        $this->setAutomaticStyle($text, $tagName, 1, $row_number);
        $this->setCellValue($this->sheet, 1, $row_number, $text);
        if(!isset($this->oneCellRowsBySheetIndex[$this->currentSheetIndex]))
            $this->oneCellRowsBySheetIndex[$this->currentSheetIndex] = array();
        $this->oneCellRowsBySheetIndex[$this->currentSheetIndex][] = $row_number;
        $this->rowHeightBySheetIndex[$this->currentSheetIndex][$row_number] = $this->getHightFromTextAndTagName($text, $tagName);
        $this->setWrapText($this->currentSheetIndex, 1, $rowNumber);
    }

    protected function generate_one_column_text_from_node($oneColumnTextNode){
        $rowNumber = $this->next_row_number();
        $this->setAutomaticStyleFromNode($oneColumnTextNode, 1, $rowNumber);
        $text = preg_replace('/[\t ][\t ]+/',' ', str_replace("\t",' ',$oneColumnTextNode->textContent));
        if(strpos($text, "\n")){
            $text = trim(str_replace("\n ","\n", $text));
            $this->setWrapText($this->currentSheetIndex, 1, $rowNumber);
            $lines = (1+substr_count($text, "\n"));
        } else $lines = 1;
        $this->setCellValue($this->sheet, 1, $rowNumber, $text);
        if(!isset($this->oneCellRowsBySheetIndex[$this->currentSheetIndex]))
            $this->oneCellRowsBySheetIndex[$this->currentSheetIndex] = array();
        $this->oneCellRowsBySheetIndex[$this->currentSheetIndex][] = $rowNumber;
        $this->rowHeightBySheetIndex[$this->currentSheetIndex][$rowNumber] = $this->getFontHeightFromTagName(
            $oneColumnTextNode->tagName) * $lines;
    }

    static function getConfig($arguments = null) {
        $file_name = __DIR__.'/config.yml';
        if(!file_exists($file_name) && file_exists("$file_name.example"))
            copy("$file_name.example", $file_name);
        $arguments = is_array($arguments) ? $arguments : explode('.',$arguments);
        $node = yaml_parse(file_get_contents($file_name));
        foreach ($arguments as $key) {
            if (isset($node[$key])) { $node = $node[$key]; }
            else return null;
        }
        return $node;
    }

    protected function handleTab($tabNode){
        $title = $tabNode->hasAttribute('title') ? $tabNode->getAttribute('title') : 'Dados';
        if($this->hasWrittenFirstLineBySheetIndex[$this->currentSheetIndex]){
            $this->sheet = $this->objPHPExcel->createSheet();
            $this->hasWrittenFirstLineBySheetIndex[$this->currentSheetIndex] = false;
            $this->currentSheetIndex++;
        }
        $this->sheet->setTitle($title);
    }

    public function mergeCells($sheetIndex, $initialColumnNumber, $initialRowNumber, $finalColumnNumber,
        $finalRowNumber) {
        if($initialColumnNumber == $finalColumnNumber && $initialRowNumber == $finalRowNumber){ return; }
        $sheet = $this->objPHPExcel->getSheet($sheetIndex);
        $mergedRanges = $sheet->getMergeCells();
        $minCellWidth = $this->stringWidthsMatrixBySheetIndex[$sheetIndex][$initialColumnNumber][$initialRowNumber];
        if(!empty($mergedRanges)) {
            foreach ($mergedRanges as $mergedRange) {
                $mergedRangeParams = self::stringRangeToSeparated($mergedRange);
                if($mergedRangeParams[0] > $finalColumnNumber) { continue; }
                if($mergedRangeParams[2] < $initialColumnNumber) { continue; }
                if($mergedRangeParams[1] > $finalRowNumber) { continue; }
                if($mergedRangeParams[3] < $initialRowNumber) { continue; }
                $this->unmergeCells($sheetIndex, $mergedRangeParams[0], $mergedRangeParams[1], $mergedRangeParams[2],
                    $mergedRangeParams[3]);
                if($initialColumnNumber != $finalColumnNumber)
                    $minCellWidth = $minCellWidth * ($finalColumnNumber-$initialColumnNumber+1);
                $initialColumnNumber = min($initialColumnNumber, $mergedRangeParams[0]);
                $initialRowNumber = min($initialRowNumber, $mergedRangeParams[1]);
                $finalColumnNumber = max($finalColumnNumber, $mergedRangeParams[2]);
                $finalRowNumber = max($finalRowNumber, $mergedRangeParams[3]);
            }
        }
        $minCellWidth = $minCellWidth/($finalColumnNumber-$initialColumnNumber+1);
        $cellStyle = $this->getCellStyle($sheet, $initialColumnNumber, $initialRowNumber);
        for ($col=$initialColumnNumber; $col <= $finalColumnNumber; $col++)
            for ($row=$initialRowNumber; $row <= $finalRowNumber; $row++){
                $this->stringWidthsMatrixBySheetIndex[$sheetIndex][$col][$row] = $minCellWidth;
                $sheet->duplicateStyle($cellStyle,self::get_letters($col).$row);
            }
        $sheet->mergeCells(self::get_letters($initialColumnNumber).$initialRowNumber.':'
            .self::get_letters($finalColumnNumber).$finalRowNumber);
    }

    public function unmergeCells($sheetIndex, $initialColumnNumber, $initialRowNumber, $finalColumnNumber,
        $finalRowNumber){
        $sheet = $this->objPHPExcel->getSheet($sheetIndex);
        $sheet->unmergeCells(self::separatedRangeToString($initialColumnNumber, $initialRowNumber, $finalColumnNumber,
            $finalRowNumber));
    }

    public static function separatedRangeToString($initialColumnNumber, $initialRowNumber, $finalColumnNumber,
        $finalRowNumber){
        return self::get_letters($initialColumnNumber).$initialRowNumber.':'.self::get_letters($finalColumnNumber)
            .$finalRowNumber;
    }

    protected static function removeInternalTags($node, array $tagNames = array('style')){
        if(!$node->hasChildNodes()) { return $node; }
        $node = $node->cloneNode(true);
        foreach ($tagNames as $tagName)
            foreach ($node->getElementsByTagname($tagName) as $foundElement)
                $foundElement->parentNode->removeChild($foundElement);
        return $node;
    }

    public static function stringRangeToSeparated($stringRange){
        if(!preg_match('/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/', $stringRange, $matches))
            return array(null,null,null,null);
        array_shift($matches);
        $matches[0] = self::get_column_number($matches[0]);
        $matches[2] = self::get_column_number($matches[2]);
        return $matches;
    }

    public function getAllSheets(){
        return $this->objPHPExcel->getAllSheets();
    }

    protected function hasStyle($node){
        return is_a($node, 'DOMElement') && !empty($this->styles[$node->tagName]);
    }

    protected function convertFontSizeToHeight($fontSize){
        // var_dump(imagettfbbox())
        return $fontSize*1.875;
    }

    protected function getFontHeightFromTagName($tagName){
        return $this->convertFontSizeToHeight($this->getFontSizeFromTagName($tagName));
    }

    protected function getIdealHeightForCell($sheetIndex, $columnNumber, $rowNumber){
        // $sheet = $this->objPHPExcel->getSheet($sheetIndex);
        //     $columnNumber = 3; //deebug;
        // $text = $this->getCellValue($sheet, $columnNumber, $rowNumber);
        // $text = '12345678901234567890';
        // // var_dump($text);
        // $font = $sheet->getStyle(self::get_letters($columnNumber).$rowNumber)->getFont();
        // if(empty($this->columnWidthBySheetIndex[$sheetIndex][$columnNumber])){
        //     $bbox = imagettfbbox($font->getSize(), 0, $font->getName(), $text);
        //     $linesCounter = (1+substr_count($text, "\n"));
        //     $height = abs($bbox[7] - $bbox[1]);
        // } else {
        //     $columnWidth = $this->columnWidthBySheetIndex[$sheetIndex][$columnNumber]-2;
        //     var_dump($columnWidth); //debug
        //     $lines = explode("\n", $text);
        //     $height = 0;
        //     $linesCounter = 0;
        //     echo'$columnWidth';var_dump($columnWidth);
        //     foreach ($lines as $line){
        //         $words = explode(" ", $line);
        //         $linesCounter++;
        //         $occupiedSpace = 0;
        //         foreach ($words as $word) {
        //             $bbox = imagettfbbox($font->getSize(), 0, $font->getName(), " $word");
        //             $occupiedSpace += $this->convertSizeUnitFromPX(abs($bbox[2] - $bbox[0]));
        //             echo'$occupiedSpace';var_dump($occupiedSpace);
        //             if($occupiedSpace >= $columnWidth){
        //                 echo'$word';var_dump($word);
        //                 $bbox = imagettfbbox($font->getSize(), 0, $font->getName(), $occupiedSpace);
        //                 $height += abs($bbox[7] - $bbox[1]);
        //                 $occupiedSpace = $this->convertSizeUnitFromPX(abs($bbox[2] - $bbox[0]));
        //                 echo'$height';var_dump($height);
        //                 echo "<br><br>";
        //                 $linesCounter++;
        //             }
        //         }
        //         $height += abs($bbox[7] - $bbox[1]);
        //         echo'height';var_dump($height);
        //         echo "<br><br>";
        //     }
        // }
        // var_dump($height*1.138);die;
        // die;
        // return $height*1.138;
    }

    protected function getFontSizeFromTagName($tagName){
        if(!empty($this->styles[$tagName]) && isset($this->styles[$tagName]['font'])
             && isset($this->styles[$tagName]['font']['size']))
            return $this->styles[$tagName]['font']['size'];
        return 8; //default
    }

    protected function duplicateStyle($originSheet, $originColumnNumber, $originRowNumber,
        $destinySheet, $destinyColumnNumber, $destinyRowNumber) {
        $destinySheet->duplicateStyle($originSheet->getStyle(
            self::get_letters($originColumnNumber).$originRowNumber),
            self::get_letters($destinyColumnNumber).$destinyRowNumber);
    }

    static protected function parseCSS($css){
        $results = array();
        preg_match_all('/(.+?)\s?\{\s?(.+?)\s?\}/', $css, $matches);
        foreach($matches[0] AS $i=>$original)
            foreach(explode(';', $matches[2][$i]) AS $attr)
                if (strlen(trim($attr)) > 0) {
                    $new_index = count($results);
                    list($name, $value) = explode(':', $attr);
                    $results[$new_index] = array(trim($matches[1][$i]), trim($name), trim($value));
                }
        return $results;
    }
    static protected function parseCSSAttributes($cssAttributes){
        $results = array();
        if(empty($cssAttributes)) { return $results; }
        foreach(explode(';', $cssAttributes) AS $attr){
            $attr = trim($attr);
            if ($attr !== '' && $attr !== "'") {
                list($name, $value) = explode(':', $attr);
                $results[$name] = trim($value, " \n\r\"");
            }
        }
        return $results;
    }

    protected function convertColorToRGB($color){
        if(strpos($color, '#') === 0){
            if(strlen($color) === 7)
                return substr($color, 1);
            if(strlen($color) === 4)
                return str_repeat(substr($color, 1, 1),2)
                . str_repeat(substr($color, 2, 1),2)
                . str_repeat(substr($color, 3, 1),2);
        }
        switch ($color) {
            case 'yellow': return 'FFFF00';
            case 'green': return '008000';
            case 'red': return 'FF0000';
            case 'orange': return 'FFA500';
            case 'gray': return '808080';
        }
    }

    public function applyCSSToDefaultStyle($cssAttributes, array $styles, $tag_name){
        if(is_string($cssAttributes))
            $cssAttributes = $this->parseCSSAttributes($cssAttributes);
        foreach ($cssAttributes as $attr => $value) {
            switch ($attr) {
                case 'border':
                    $styles['borders'] = array('allborders' => array('style' => $value));
                    break;
                case 'background-color':
                    $styles['fill'] = array('type' => PHPExcel_Style_Fill::FILL_SOLID,
                        'color' => array('rgb' => $this->convertColorToRGB($value)));
                    break;
                case 'number-format':
                    $styles['numberformat'] = array('code' => $value);
                case 'text-align':
                    $styles['alignment'] = array('horizontal' => $value);
            }
        }
        return $styles;
    }

    static protected function exitIfInvalidHTML($html){
        $lastError = error_get_last();
        if(empty($lastError)) { return; }
        if(in_array($lastError['type'] , array(E_ERROR, E_PARSE, E_CORE_ERROR,
            E_COMPILE_ERROR, E_USER_ERROR, E_RECOVERABLE_ERROR, E_DEPRECATED,
            E_USER_DEPRECATED)))
            exit("Error found: {$lastError['message']} on {$lastError['file']} at line {$lastError['line']}.$html");
    }
}
