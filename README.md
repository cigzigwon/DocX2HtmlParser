DocX2HtmlParser
===============

The best MS Word docx parser to HTML out and more to come...
Support for .odt (Open Document Text) coming soon.

Usage
-----
$docxParser = new DocX2HtmlParser();
$docxParser->setFile('~/path/to/file');
$html = $doc->toHtml();
echo $html;
