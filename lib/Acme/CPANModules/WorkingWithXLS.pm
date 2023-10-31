package Acme::CPANModules::WorkingWithXLS;

use strict;

use Acme::CPANModulesUtil::Misc;

# AUTHORITY
# DATE
# DIST
# VERSION

my $text = <<'_';

The following are tools (programs, modules, scripts) to work with Excel formats
(XLS, XLSX) or other spreadsheet formats like LibreOffice Calc (ODS).



**Parsing**

<pm:Spreadsheet::Read> is a common-interface front-end for
<pm:Spreadsheet::ReadSXC> (for reading LibreOffice Calc ODS format) or one of
<pm:Spreadsheet::ParseExcel>, <pm:Spreadsheet::ParseXLSX>, or Spreadsheet::XLSX
(for reading XLS or XLSX, although Spreadsheet::XLSX is strongly discouraged
because it is a quick-and-dirty hack). Spreadsheet::Read can also read CSV via
Text::CSV_XS. The module can return information about cell's attributes
(formatting, alignment, and so on), merged cells, etc. The distribution of this
module also comes with some CLIs like <prog:xlscat>, <prog:xlsx2csv>.

<pm:Data::XLSX::Parser> which claims to be a "faster XLSX parser". Haven't used
this one personally or benchmarked it though.


**Getting information**

<pm:Spreadsheet::Read>

<prog:xls-info> from <pm:App::XLSUtils>



**Iterating/processing with Perl code**

<prog:XLSperl> CLI from <pm:App::XLSperl> lets you iterate each cell (with
'XLSperl -ne' or row with 'XLSperl -ane') with a Perl code, just like you would
each line of text with `perl -ne` (in fact, the command-line options of XLSperl
mirror those of perl). Only supports the old Excel format (XLS not XLSX). Does
not support LibreOffice Calc format (ODS). If you feed it unsupported format, it
will fallback to text iterating, so if you feed it XLSX or ODS you will iterate
chunks of raw binary data.

<prog:xls-each-cell> from <pm:App::XLSUtils>

<prog:xls-each-row> from <pm:App::XLSUtils>



**Converting to CSV**

<prog:xlsx2csv> from <pm:Spreadsheet::Read>. Since it's based on
Spreadsheet::Read, it can read XLS/XLSX/ODS. It always outputs to file and not
to stdout.

`CATDOC` (<http://www.wagner.pp.ru/~vitus/software/catdoc/>) contains following
the programs `catdoc` (to print the plain text of Microsoft Word documents to
standard output), <prog:xls2csv> (to convert Microsoft Excel workbook files to
CSV), and `catppt` (to print plain text of Mirosoft PowerPoint presentations to
standard output). Available as Debian package. They only support the older
format (XLS and not XLSX). They do not support LibreOffice Calc format (ODS).

<prog:xls2csv> from <pm:App::XLSUtils>


**Generating XLS**

TBD

_

our $LIST = {
    summary => 'List of modules to work with Excel formats (XLS, XLSX) or other spreadsheet formats like LibreOffice Calc (ODS)',
    description => $text,
    tags => ['task'],
};

Acme::CPANModulesUtil::Misc::populate_entries_from_module_links_in_description;

1;
# ABSTRACT:

=head1 prepend:SEE ALSO

L<Acme::CPANModules::WorkingWithCSV>
