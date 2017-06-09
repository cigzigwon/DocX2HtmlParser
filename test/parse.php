<?php

require(dirname(__FILE__) . '/../vendor/autoload.php');

$file = '~/path/to/file';
$docx = new DocX2HtmlParser();
$docx->setFile($file);
$html = $docx->toHtml();
echo $html;
