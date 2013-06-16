#!/bin/bin/perl
#
# Copyright (C) 2013 by Lieven Hollevoet

# This test runs basic module tests

use strict;
use Test::More;

BEGIN { use_ok 'App::CPRReporter'; }
BEGIN { use_ok 'Test::Exception'; }

# Check we get an error message on missing input parameters
my $reporter;

can_ok ('App::CPRReporter', qw(employees certificates run));

throws_ok { $reporter = App::CPRReporter->new() } qr/Attribute .+ is required/, "Checking missing parameters";
throws_ok { $reporter = App::CPRReporter->new(employees => 't/stim/missing_file.xlsx', certificates => 't/stim/missing_file.xml') } qr/File does not exist.+/, "Checking missing xml file";

$reporter = App::CPRReporter->new(employees => 't/stim/employees.xlsx', certificates => 't/stim/certificates.xml');
ok $reporter, 'object created';
ok $reporter->isa('App::CPRReporter'), 'and it is the right class';

$reporter->run();


done_testing();