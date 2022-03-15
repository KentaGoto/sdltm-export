use strict;
use warnings;
use utf8;
use DBI;

my $dbh = DBI->connect("dbi:SQLite:dbname=test.sdltm");

# ソースとターゲットのみ対象とする
my $select = "select source_segment, target_segment from translation_units;"; 

my $sth = $dbh->prepare($select);
$sth->execute;

open my $out, '>', 'result.txt' or die;

while(my ($source, $target) = $sth->fetchrow()){
	$source = &seikei($source);
	$target = &seikei($target);
	print {$out} $source . "\t" . $target . "\n";
}

$dbh->disconnect;

close $out;
print "Done!\n";

sub seikei {
	my $s = shift;
	$s =~ s{^.+<Value>(.+)</Value>.+$}{$1};

	return $s;
}
