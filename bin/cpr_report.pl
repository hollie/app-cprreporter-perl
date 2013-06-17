#! /usr/bin/env perl

use strict;
use 5.012;
use App::CPRReporter;

my $reporter = App::CPRReporter->new(employees => 'employees.xlsx', certificates => 'certificates.xml');
$reporter->run();

# PODNAME: cpr_reporter.pl
# ABSTRACT: Example application
