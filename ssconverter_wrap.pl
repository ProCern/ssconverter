#!/usr/bin/perl
use File::Basename;

# Configuration

my $WORKDIR_LOCAL = "/pcshare/ssconverter";
my $WORKDIR_REMOTE = "/pcshare/ssconverter";

#Mangle the filenames a bit
$FULL_FROM = $ARGV[0];
$FULL_TO = $ARGV[1];
$SFROM = basename($FULL_FROM);
$STO = basename($FULL_TO);

# Create the temp file to operate on
system("cp $FULL_FROM $WORKDIR_LOCAL/$SFROM");

# run the remote conversion utility
system("ssh faxserver /usr/local/bin/ssconverter.pl $WORKDIR_REMOTE/$SFROM $WORKDIR_REMOTE/$STO")
