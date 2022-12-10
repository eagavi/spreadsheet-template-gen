#!/usr/bin/perl

use strict;
use warnings;

use Excel::Writer::XLSX;

my $workbook = Excel::Writer::XLSX->new('supplier_list_template.xlsx');

my $worksheet = $workbook->add_worksheet();
$worksheet->keep_leading_zeros(1);

my $center_alignment = $workbook->add_format(align=>'center');

$worksheet->write('A1', 'Supplier', $center_alignment);
$worksheet->write('B1', 'Invoice No', $center_alignment);
$worksheet->write('C1', 'Net Value', $center_alignment);
$worksheet->write('D1', 'Added VAT', $center_alignment);
$worksheet->write('E1', 'Net Value & VAT', $center_alignment);

$worksheet->write('C34', '=SUM(C2:C34)', $center_alignment);
$worksheet->write('E34', '=SUM(E2:E34)', $center_alignment);

$workbook->close();


