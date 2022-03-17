use strict;
use warnings;
use utf8;
use DBI;
use Win32::OLE qw( in with CP_UTF8 );
use Win32::OLE::Const 'Microsoft Excel';
use File::Basename;
use Encode;
use File::Find::Rule;

print "Dir: ";
chomp( my $dir = <STDIN> );
$dir =~ s{^"}{};
$dir =~ s{"$}{};
my @sdltm = File::Find::Rule->file->name(qr/\.sdltm$/i)->in($dir);

my $dirname = dirname (__FILE__);

Win32::OLE->Option( CP => CP_UTF8 );
my $Excel = Win32::OLE->new( 'Excel.Application', 'Quit' );
$Excel->{'Visible'}       = 0;
$Excel->{'DisplayAlarts'} = 0;

print "Processing...\n";

foreach my $sdltm ( @sdltm ){
	print $sdltm . "\n";

	my $book  = $Excel->Workbooks->add();
	my $sheet = $book->Worksheets(1);
	
	$sheet->Range("A1")->{'Value'} = 'Source';
	$sheet->Range("B1")->{'Value'} = 'Target';
	
	my $dbh = DBI->connect("dbi:SQLite:dbname=$sdltm");
	
	# source and target only
	my $select = "select source_segment, target_segment from translation_units;"; 
	
	my $sth = $dbh->prepare($select);
	$sth->execute;
	
	my $count = 2;
	while(my ($source, $target) = $sth->fetchrow()){
		$source = &seikei($source);
		$target = &seikei($target);
		$sheet->Range("A$count")->{'Value'} = $source;
		$sheet->Range("B$count")->{'Value'} = $target;
		$count++;
	}
	
	$dbh->disconnect;

	( my $xlsx = $sdltm ) =~ s{\.sdltm}{.xlsx}i;
	
	$book->SaveAs( $xlsx );
	$book->Close;
}

$Excel->quit();

print "\nDone!\n";

sub seikei {
	my $s = shift;

	$s = decode('utf8', $s); 

	my $str;

	# Valueタグ内のテキストを取る。複数ある場合に対応
	while ( $s =~ s{<Value>(.+?)</Value>}{$1}s ) {
		$str .= $1;
	}

	return $str;
}
