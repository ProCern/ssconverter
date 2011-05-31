#!/usr/bin/perl

if ( $#ARGV < 1 || $ARGV[0] =~ /^\-h$/ ) {
  print "Usage:\n";
  print "ssconverter.pl <inputfile.csv> <outputfile.xls>\n";
  print "Generates EXCEL 2003-2007 format xls binary files\n";
}

else {
  $INFILE = $ARGV[0];
  $OUTFILE = $ARGV[1];
}

unless (open INF, "<$INFILE" ) {
  die "Error: cannot open file $INFILE\n";
}

my $l = 0;
my @widths;
my @content;
my @header;
my $cols=0;

# Parse the incoming file
while ( <INF> ) {
  print "$l\n";
  chomp;
  my $bareline = $_;
  my @line = split(/,/);
  $cols = $#line;
  foreach $i (0..@line) {
    if ($l == 0) { # processing the header line
      $width[$i] = length($line[$i]);
      $header[$i] = $line[$i];
    }
    else {
      # check for delimiter / parsing error
      if ( $#line != $cols ) {
        print STDERR "Error: parsing problem - number of cols does not matcher header on line: $bareline\n";
        next;
      }
      # ensure the width is the widest width
      if ( length($line[$i]) > $width[$i] ) {
        $width[$i] = length($line[$i]);
      }
      if ( $line[$i] =~ /^"(.*)"$/ ) {
        $content[$l-1][$i] = $1;
      }
      else {
        $content[$l-1][$i] = $line[$i];
      }
    }
  }
  ++$l;
}
close(INF);

# Setup the spreadsheet, write header and column widths

for $i (0..$#header) {
  print "$header[$i]:$width[$i] "
}
print "\n";

# Write the data to the spreadsheet

foreach $row (0..@content) {
  print "$row:\n";
  foreach $col (0..@{$content[$row]}) {
    print "$col: $content[$row][$col] ";
  }
  print "\n\n";
}
