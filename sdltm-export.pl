use strict;
#use warnings;
use utf8;
use DBI;

my $dbh = DBI->connect("dbi:SQLite:dbname=test.sdltm");
my $select = "select * from translation_units;";
my $sth = $dbh->prepare($select);
$sth->execute;

open my $out, '>', 'result.txt' or die;

while(my @row = $sth->fetchrow_array){
	#print {$out} join("\t", @row), "\n";
	print {$out} join(", ", @row), "\n";
}

#$sth->finish;
#undef $sth;

$dbh->disconnect;

close $out;
