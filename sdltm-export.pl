use strict;
use warnings;
use utf8;
use DBI;
use Win32::OLE qw( in with CP_UTF8 );
use Win32::OLE::Const 'Microsoft Excel';
use File::Basename;
use Encode;

my $dirname = dirname (__FILE__);

Win32::OLE->Option( CP => CP_UTF8 );
my $Excel = Win32::OLE->new( 'Excel.Application', 'Quit' );
$Excel->{'Visible'}       = 0;
$Excel->{'DisplayAlarts'} = 0;

my $book  = $Excel->Workbooks->add();
my $sheet = $book->Worksheets(1);

$sheet->Range("A1")->{'Value'} = 'Source';
$sheet->Range("B1")->{'Value'} = 'Target';

my $dbh = DBI->connect("dbi:SQLite:dbname=test.sdltm");

# ソースとターゲットのみ対象とする
my $select = "select source_segment, target_segment from translation_units;"; 

my $sth = $dbh->prepare($select);
$sth->execute;


#open my $out, '>', 'result.txt' or die;

my $count = 2;
while(my ($source, $target) = $sth->fetchrow()){
	$source = &seikei($source);
	$target = &seikei($target);
	#print {$out} $source . "\t" . $target . "\n";
	$sheet->Range("A$count")->{'Value'} = $source;
    $sheet->Range("B$count")->{'Value'} = $target;
    $count++;
}

$dbh->disconnect;

#close $out;
$book->SaveAs( $dirname . '\\' . 'results.xlsx' );
$book->Close;
$Excel->quit();

print "Done!\n";

sub seikei {
	my $s = shift;
	$s = decode('utf8', $s);
	$s =~ s{^.+<Value>(.+)</Value>.+$}{$1};

	return $s;
}
