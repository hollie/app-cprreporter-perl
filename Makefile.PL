
use strict;
use warnings;

use 5.012;

use ExtUtils::MakeMaker 6.30;



my %WriteMakefileArgs = (
  "ABSTRACT" => "Application to merge various datasets info an overview of who followed CPR training",
  "AUTHOR" => "Lieven Hollevoet <hollie\@cpan.org>",
  "BUILD_REQUIRES" => {
    "Test::Exception" => 0,
    "Test::More" => 0,
    "Test::Warn" => 0
  },
  "CONFIGURE_REQUIRES" => {
    "ExtUtils::MakeMaker" => "6.30"
  },
  "DISTNAME" => "App-CPRReporter",
  "EXE_FILES" => [
    "bin/cpr_report.pl"
  ],
  "LICENSE" => "perl",
  "NAME" => "App::CPRReporter",
  "PREREQ_PM" => {
    "Carp" => 0,
    "Data::Dumper" => 0,
    "Getopt::Long" => 0,
    "Moose" => 0,
    "Pod::Usage" => 0,
    "Spreadsheet::XLSX" => 0,
    "Text::Fuzzy::PP" => 0,
    "Text::Iconv" => 0,
    "Text::ResusciAnneparser" => 0,
    "autodie" => 0,
    "namespace::autoclean" => 0,
    "strict" => 0,
    "warnings" => 0
  },
  "VERSION" => "0.01",
  "test" => {
    "TESTS" => "t/*.t"
  }
);


unless ( eval { ExtUtils::MakeMaker->VERSION(6.56) } ) {
  my $br = delete $WriteMakefileArgs{BUILD_REQUIRES};
  my $pp = $WriteMakefileArgs{PREREQ_PM};
  for my $mod ( keys %$br ) {
    if ( exists $pp->{$mod} ) {
      $pp->{$mod} = $br->{$mod} if $br->{$mod} > $pp->{$mod};
    }
    else {
      $pp->{$mod} = $br->{$mod};
    }
  }
}

delete $WriteMakefileArgs{CONFIGURE_REQUIRES}
  unless eval { ExtUtils::MakeMaker->VERSION(6.52) };

WriteMakefile(%WriteMakefileArgs);



