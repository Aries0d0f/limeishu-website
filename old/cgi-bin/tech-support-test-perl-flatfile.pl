#!/usr/local/bin/perl

use CGI;
use CGI::Carp qw(fatalsToBrowser);

$q = new CGI;
print $q->header();

print "This test script has been uploaded to your site by web hosting technical support because of a support request (IDS-8022157) for your hosting account.<br><br>";
print "If this test script executes correctly you will see the current time and date printed below.  this information will also be written to a file.  Every time you refresh this page, you will see the updated time and date and you will also see the full contents of the flat file we have been writing our timestamp too.  Otherwise you will see an error describing the failure.<br><br>";
print "<hr>\n";
print "<br>\n";


## Need to set the current time and date
$date = localtime();


## Open the file for writing.  We want to append to the end of the file.
## It will be created if it does not already exist.
open(DATE, ">>../data/tech-support-test-perl-flatfile.txt") || die "Unable to open flat file for writing: $!<br>\n";


## Write the current date to the end of the file
print DATE "$date\n";
print "Wrote \"$date\" to flat file.<br>\n";

## Close the file handle
close(DATE);


## Print the separator
print("<br>  <hr>  <br>\n");
print("The following are the contents of the file we just wrote our timestamp to.  As you refresh you will see additional timestamps listed.<br><br>\n");


## Open the file for reading
open(FILE, "../data/tech-support-test-perl-flatfile.txt") || die "Unable to open flat file for reading: $!<br>\n";

## Display the contents of the file
foreach(<FILE>){
	print "$_<br>\n";
}

## Close the filehandle
close(FILE);