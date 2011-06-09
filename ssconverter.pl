#!/usr/bin/perl

#use strict;
use Spreadsheet::WriteExcel;

my $DEBUG = 0;

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
  if ($DEBUG) {print "$l:\n";}
  s/\r//; # deal with pesky \r\n from windows
  chomp($_);
  my $bareline = $_;
  my @line;
  if ( scalar( () = $bareline =~ /\t/g ) > 2 ) {
    @line = split(/\t/,$bareline);
    if ($DEBUG) { print "Selected TSV format\n"; }
  }
  elsif ( scalar( () = $bareline =~ /\",\"/g ) > 2 ) {
    @line = split(/\",\"/,$bareline);
    if ($DEBUG) { print "Selected Quote Enclosed CSV format\n"; }
  }
  elsif ( scalar( () = $bareline =~ /,/g ) > 2 ) {
    @line = split(/,/,$bareline);
    if ($DEBUG) { print "Selected CSV format\n"; }
  }
  else {
    print STDERR "line not formatted as csv or tsv: $bareline\n";
  }
  # Loop through all of the elements and remove all " characters
  for my $i (0..$#line) {
    $line[$i] =~ s/\"//g;
  }
  if ($l == 0 ) { # handle the header differently than the body
    $cols = $#line;
    for my $i (0..$#line) {
      $width[$i] = length($line[$i]);
      $header[$i] = $line[$i];
      if ($DEBUG) {print "set header[$i] = $line[$i] width: $width[$i]\n";}
    }
  }
  else { # handle the body/data
    if ( $#line != $cols ) {
      print STDERR "Error: parsing problem - number of cols does not matcher header on line: $bareline\n";
      next;
    }
    for my $i (0..$#line) {
      # check for delimiter / parsing error
      # ensure the width is the widest width
      if ( length($line[$i]) > $width[$i] ) {
        $width[$i] = length($line[$i]);
      }
      $content[$l-1][$i] = $line[$i];
    }
  }
  ++$l;
}
close(INF);

# Setup the spreadsheet, write header and column widths
if ($DEBUG) {print "########################\n";}

my $workbook = Spreadsheet::WriteExcel->new($OUTFILE);
my $worksheet = $workbook->add_worksheet();
my $strfmt = $workbook->add_format(num_format => '@');

for my $i (0..$#header) {
  if ($DEBUG) {print "$header[$i]:$width[$i] ";}
  if ( $header[$i] =~ /acct_no/i ) {
    if ($DEBUG) {print " in special header\n";}
    $worksheet->set_column($i,$i,int($width[$i]*1.1),$strfmt);
  }
  else {
    $worksheet->set_column($i,$i,int(($width[$i]+1)*1.2));
  }
  $worksheet->write_string(0,$i,$header[$i]);
}

if ($DEBUG) {print "\n";}



# Write the data to the spreadsheet

foreach $row (0..@content) {
  if ($DEBUG) {print "$row:\n";}
  foreach $col (0..@{$content[$row]}) {
    if ($DEBUG) {print "$col: $content[$row][$col] ";}
    if ( $content[$row][$col] =~ /\d{10}/ ) {
      $worksheet->write_string($row+1,$col,$content[$row][$col]);
    }
    elsif ( $content[$row][$col] =~ /^$/ ) {
      if ($DEBUG) {print " in blank handller\n";}
      #$worksheet->write_blank($row,$col);
    }
    else {
      $worksheet->write($row+1,$col,$content[$row][$col]);
    }
  }
  if ($DEBUG) {print "\n\n";}
}
