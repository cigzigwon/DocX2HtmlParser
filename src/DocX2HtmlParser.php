<?php
class DocX2HtmlParser {
  
  /**
   * toggles debug output for development reasons
   *
   * @var bool
   */
  private $debugMode;
  
  /**
   * internal docx data stream to read
   *
   * @var string
   */
  private $fileData;
  
  /**
   * the zip archive data for access to xml data
   *
   * @var ZipArchive
   */
  private $zip;
  
  /**
   * array to house errors
   *
   * @var array
   */
  private $errors = [];
  
  /**
   * array to house styles
   *
   * @var array
   */
  private $styles = [];
  
  /**
   * class instance constructor
   *
   * @param boolean $debugMode toggles debug mode on/off
   */
  public function __construct($debugMode = FALSE) {
    $this->debugMode = $debugMode;
  }
  
  /**
   * loads the file by using ZipArchive to return data for parsing
   *
   * @param  string $file the file path to load
   * @return string       the resulting xml markup from file
   */
  private function load($file) {
    
    if (file_exists($file)) {
      $this->zip = new ZipArchive();
      
      if (TRUE === $this
        ->zip
        ->open($file)) {
        
        if (FALSE !== ($styleIndex = $this
          ->zip
          ->locateName('word/styles.xml'))) {
          $stylesXml = $this
            ->zip
            ->getFromIndex($styleIndex);
          
          $xml = simplexml_load_string($stylesXml);
          $namespaces = $xml->getNamespaces(TRUE);
          $children = $xml->children($namespaces['w']);
          
          foreach ($children
            ->style as $s) {
            $attr = $s->attributes('w', TRUE);
            
            if (isset($attr['styleId'])) {
              $tags = [];
              $attrs = [];
              
              foreach (get_object_vars($s
                ->rPr) as $tag => $style) {
                $att = $style->attributes('w', TRUE);
                
                switch ($tag) {
                  case "b":
                    $tags[] = 'strong';
                  break;
                  case "i":
                    $tags[] = 'em';
                  break;
                  case "u":
                    $tags[] = 'u';
                  break;
                  case "color":
                    $attrs[] = 'color:#' . $att['val'];
                  break;
                  case "sz":
                    $attrs[] = 'font-size:' . $att['val'] / 2 . 'pt';
                  break;
                }
              }
              
              $styles[(string)$attr['styleId']] = array(
                'tags' => $tags,
                'attrs' => $attrs,
              );
            }
          }
          
          $this->styles = $styles;
        }
        
        if (FALSE !== ($index = $this
          ->zip
          ->locateName('word/document.xml'))) {
          
          $data = $this
            ->zip
            ->getFromIndex($index);
          
          return $data;
        }
      } else {
        $this->errors[] = 'Could not open file.';
      }
    } else {
      $this->errors[] = 'File does not exist.';
    }
  }
  
  /**
   * sets the file path for ease of use
   *
   * @param string $path the location of the file
   */
  public function setFile($path) {
    $this
      ->fileData = $this->load($path);
  }
  
  /**
   * outputs only inline text of the xml markup
   *
   * @return mixed
   */
  public function toPlainText() {
    if ($this
      ->fileData) {
      return strip_tags($this->fileData);
    } else {
      return FALSE;
    }
    
    $this
      ->zip
      ->close();
  }
  
  /**
   * outputs html markup from parsing xml
   *
   * @return string   resulting html markup
   */
  public function toHtml() {
    if ($this->fileData) {
      $html = '';
      $xml = simplexml_load_string($this->fileData);
      $namespaces = $xml->getNamespaces(TRUE);
      $children = $xml->children($namespaces['w']);
      
      if ($this->debugMode) {
        $html = '<!doctype html><html><head><meta http-equiv="Content-Type" content="text/html;charset=utf-8" /><title></title><style>span.block { display: block; }</style></head><body>';
      }
      
      $openList = TRUE;
      $closeList = FALSE;
      
      foreach ($children->body as $body) {
        
        foreach ($body as $prop => $elem) {
          $startTags = [];
          $startAttrs = [];

          if ('p' === $prop) {
            $isHeading = FALSE;
            $style = '';
            
            if ($elem
              ->pPr
              ->pStyle) {
              $objectAttrs = $elem
                ->pPr
                ->pStyle
                ->attributes('w', TRUE);
              
              $objectStyle = (string)$objectAttrs['val'];

              if (preg_match('/Heading/', $objectStyle)) $isHeading = TRUE;
              
              if (isset($this
                ->styles[$objectStyle])) {
                $startTags = $this->styles[$objectStyle]['tags'];
                $startAttrs = $this->styles[$objectStyle]['attrs'];
              }
            }
            
            if ($elem
              ->pPr
              ->spacing) {
              $att = $elem
                ->pPr
                ->spacing
                ->attributes('w', TRUE);
              
              if (isset($att['before'])) {
                $style.= 'padding-top:' . ($att['before'] / 10) . 'px;';
              }
              
              if (isset($att['after'])) {
                $style.= 'padding-bottom:' . ($att['after'] / 10) . 'px;';
              }
            }
            
            $li = FALSE;
            
            if ($elem
              ->pPr
              ->numPr && !$isHeading) {
              $li = TRUE;
              
              if ($openList) {
                $openList = FALSE;
                $closeList = TRUE;
                
                $listTags = $this
                  ->getListFormatting($elem
                  ->pPr
                  ->numPr);
                
                $html.= $listTags[0];
              }
              
              $html.= '<li>';
            } else {
              
              if (!$closeList) {
                
                if ($elem
                  ->pPr
                  ->jc) {
                  $att = $elem
                    ->pPr
                    ->jc
                    ->attributes('w', TRUE);
                  
                  $html.= '<p style="text-align:' . $att['val'] . ';">';
                } else {
                  
                  if (!$li) $html.= '<p>';
                }
              }
            }
            
            $html.= $this->parseText($elem, $startTags, $startAttrs);
            
            if ($li) {
              $html.= '</li>';
            } else {
              
              if ($closeList) {
                $openList = TRUE;
                $closeList = FALSE;
                $html.= $listTags[1];
              }
              
              $html.= "</p>";
            }
          } else if ($prop === 'tbl') {
            
            if ($closeList) {
              $openList = TRUE;
              $closeList = FALSE;
              $html.= $listTags[1];
            }
            
            $html.= '<table class="table" border="1">';
            
            foreach ($elem->tr as $tr) {
              $html.= '<tr>';
              
              foreach ($tr
                ->tc as $tc) {
                
                if ($tc
                  ->tcPr
                  ->gridSpan) {
                  $att = $tc
                    ->tcPr
                    ->gridSpan
                    ->attributes('w', TRUE);
                  
                  $html.= '<td colspan="' . $att['val'] . '">';
                } else {
                  $html.= '<td>';
                }
                
                if ($tc
                  ->p) {
                  
                  foreach ($tc->p as $p) {
                    $li = FALSE;
                    
                    if ($p
                      ->pPr
                      ->numPr) {
                      $li = TRUE;
                      $html.= '<li>';
                    }
                    
                    if ($p
                      ->pPr
                      ->jc) {
                      $att = $p
                        ->pPr
                        ->jc
                        ->attributes('w', TRUE);
                      $html.= '<p style="text-align:center">';
                    } else {
                      $html.= '<p>';
                    }
                    
                    $html.= $this->parseText($p, $startTags, $startAttrs);
                    $html.= '</p>';
                    
                    if ($li) {
                      $html.= '</li>';
                    }
                  }
                }
                
                $html.= '</td>';
              }
              
              $html.= '</tr>';
            }
            
            $html.= '</table>';
          }
        }
      }
      
      $this
        ->zip
        ->close();
      
      $regex = <<<'END'
/
  (
    (?: [\x00-\x7F]                 # single-byte sequences   0xxxxxxx
    |   [\xC0-\xDF][\x80-\xBF]      # double-byte sequences   110xxxxx 10xxxxxx
    |   [\xE0-\xEF][\x80-\xBF]{2}   # triple-byte sequences   1110xxxx 10xxxxxx * 2
    |   [\xF0-\xF7][\x80-\xBF]{3}   # quadruple-byte sequence 11110xxx 10xxxxxx * 3 
    ){1,100}                        # ...one or more times
  )
| .                                 # anything else
/x
END;
      
      preg_replace($regex, '$1', $html);
      
      if ($this->debugMode) {
        $html.= '</body></html>';
        echo $html;
        exit;
      }
      
      return $html;
    }
  }
  
  /**
   * parses most inline text and formats it as well
   *
   * @param  SimpleXMLElement $elem  parent element of xml
   * @param  array            $tags  houses array of html tags
   * @param  array            $attrs houses array of html attributes
   * @return string                  resulting html markup
   */
  private function parseText($elem, array $tags, array $attrs) {
    $html = '';
    $xmlFormats = array(
      'Strong',
      'single',
    );
    
    foreach ($elem as $p => $line) {
      $text = '';
      
      if ('r' === $p) {
        
        foreach (get_object_vars($line->pPr) as $k => $v) {
          
          if ($k === 'numPr') {
            $tags[] = 'li';
          }
        }
        
        foreach ($line->rPr as $type => $styles) {
          
          foreach (get_object_vars($styles) as $tag => $style) {
            $att = $style->attributes('w', TRUE);
            
            switch ($tag) {
              case 'rStyle':
                if (isset($att['val']) && in_array($att['val'], $xmlFormats)) {
                  $tags[] = 'strong';
                }
              break;
              case "b":
                $tags[] = 'strong';
              break;
              case "i":
                $tags[] = 'em';
              break;
              case "u":
                if (isset($att['val']) && in_array($att['val'], $xmlFormats)) {
                  $tags[] = 'u';
                }
              break;
              case "color":
                $attrs[] = 'color:#' . $att['val'];
              break;
              case "sz":
                $attrs[] = 'font-size:' . $att['val'] / 2 . 'pt';
              break;
            }
          }
        }
        
        $text = $line->t;
      } else if ('hyperlink' === $p) {
        $att = $line->attributes('r', TRUE);
        $text = $this
          ->getHyperlinkFromRel($att['id'], $line
          ->r
          ->t);
      }
      
      if (!empty($text)) {
        $openTags = '';
        $closeTags = '';
        
        foreach ($tags as $tag) {
          $openTags.= '<' . $tag . '>';
          $closeTags.= '</' . $tag . '>';
        }
        
        $html.= '<span style="' . implode(';', $attrs) . '">' . $openTags . $text . $closeTags . '</span>';
        $attrs = [];
        $tags = [];
      }
    }
    
    return $html;
  }
  
  /**
   * grabs all the hyperlinks from the related internal data file
   *
   * @param  string $id   the id of the hyperlink data
   * @param  string $text text to place inside the anchor
   * @return string       the resulting anchor html markup
   */
  private function getHyperlinkFromRel($id, $text) {
    if (FALSE !== ($index = $this
      ->zip
      ->locateName('word/_rels/document.xml.rels'))) {
      $xml = $this
        ->zip
        ->getFromIndex($index);
      
      $reader = new XMLReader;
      $reader->xml($xml);
      
      while ($reader
        ->read()) {
        
        if ('Relationship' === $reader
          ->name && $reader
          ->nodeType === XMLReader::ELEMENT) {
          
          if ($id == $reader
            ->getAttribute('Id')) {
            return '<a href="' . $reader->getAttribute('Target') . '" target="_blank">' . $text . '</a>';
          }
        }
      }
    }
    
    return FALSE;
  }
  
  /**
   * grabs formatting for list styles from a related internal data file
   *
   * @param  SimpleXMLElement $numPr the list object element
   * @return array                   wrapping tags for ordered and unordered lists
   */
  private function getListFormatting($numPr) {
    $id = $numPr
      ->numId
      ->attributes('w', TRUE) ['val'];
    
    $level = $numPr
      ->ilvl
      ->attributes('w', TRUE) ['val'];
    
    if (FALSE !== ($index = $this
      ->zip
      ->locateName('word/numbering.xml'))) {
      $xml = $this
        ->zip
        ->getFromIndex($index);
      
      $doc = new DOMDocument();
      $doc->preserveWhiteSpace = FALSE;
      $doc->loadXML($xml);
      
      $xpath = new DOMXPath($doc);
      $nodes = $xpath->query('/w:numbering/w:num[@w:numId=' . $id . ']/w:abstractNumId');
      
      if ($nodes
        ->length) {
        $id = $nodes
          ->item(0)
          ->getAttribute('w:val');
        
        $nodes = $xpath->query('/w:numbering/w:abstractNum[@w:abstractNumId=' . $id . ']/w:lvl[@w:ilvl=' . $level . ']/w:numFmt');
        
        if ($nodes
          ->length) {
          $listFormat = $nodes
            ->item(0)
            ->getAttribute('w:val');
          
          if ($listFormat === 'bullet') {
            return ['<ul>', '</ul>'];
          } else if ($listFormat === 'decimal') {
            return ['<ol>', '</ol>'];
          }
        }
      }
    }
    
    return ['<ul class="list-unstyled">', '</ul>'];
  }
  
  /**
   * fetches arrors from internal store but is not currently implemented
   *
   * @return array  an array of error data
   */
  public function getErrors() {
    return $this->errors;
  }
}
