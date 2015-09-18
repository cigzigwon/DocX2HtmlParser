DocX2HtmlParser (PHP)
=====================

The best MS Word docx parser to HTML out and more to come...

Support for .odt (Open Document Text) coming soon.

Usage
-----
$docx = new DocX2HtmlParser();

$docx->setFile('~/path/to/file');

$html = $docx->toHtml();

echo $html;

