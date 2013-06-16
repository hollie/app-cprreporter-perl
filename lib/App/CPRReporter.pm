use strict;
use warnings;
package App::CPRReporter;

use Moose;
use namespace::autoclean;
use 5.012;
use autodie;

use Carp qw/croak carp/;
use Text::ResusciAnneparser;
use Spreadsheet::XLSX;
use Text::Iconv;
use Data::Dumper;


has employees => (
    is => 'ro',
    isa => 'Str',
    required => 1,
);

has certificates => (
    is => 'ro',
    isa => 'Str',
    required => 1,
);

# Actions that need to be run after the constructor
sub BUILD {
    my $self = shift;
    # Add stuff here
    $self->{_certificates} = Text::ResusciAnneparser->new(infile => $self->{certificates});
    $self->{_employees} = $self->_employee_parser;
}

# Run the application, merging the info of the certificates and the employees
sub run {
	my $self = shift;
	
	# Certificates are here
	# TODO adapt the parser module so that the level of hierarchy in this hash is less deep
	
	my $certs = $self->{_certificates}->{_data}->{certs};
	foreach my $date (sort keys $certs) {
		foreach my $certuser (@{$certs->{$date}}){
			my $fullname = uc($certuser->{familyname}) . " " . uc($certuser->{givenname});
			say "Certificate found for $fullname";
			
			# TODO Check if certificate date is already filled in and of is it keep the most recent one.
			# Might not be required because we sort the date keys.
			if (defined $self->{_employees}->{$fullname}){
				# Fill in certificate
				$self->{_employees}->{$fullname}->{cert} = $date;
			} else {
				# Oops: user not found in personel database
				carp "Warning: employee '$fullname' not found in employee database"
			}
		}
	}
	
	my $training = $self->{_certificates}->{_data}->{training};
	foreach my $traininguser (@{$training}) {
		my $fullname = uc($traininguser->{familyname}) . " " . uc($traininguser->{givenname});
		say "Training found for $fullname";
		
			# TODO deduplicate this code with a local function, see above
			if (defined $self->{_employees}->{$fullname}){
				# Fill in training
				$self->{_employees}->{$fullname}->{cert} = 'training';
			} else {
				# Oops: user not found in personel database
				carp "Warning: employee '$fullname' not found in employee database"
			}
	}
	
	# now run the stats, for every dienst separately report
	my $stats;
	foreach my $employee (keys $self->{_employees}) {
		my $dienst = $self->{_employees}->{$employee}->{dienst};
		my $cert   = $self->{_employees}->{$employee}->{cert} || 'none';
		$stats->{employee_count} +=1;
		
		if ($cert eq 'none') {
			$stats->{$dienst}->{'not_started'}->{count} += 1;
			push (@{$stats->{$dienst}->{'not_started'}->{list}}, $employee);
		} elsif ($cert eq 'training') {
			$stats->{$dienst}->{'training'}->{count} += 1;
			push (@{$stats->{$dienst}->{'training'}->{list}}, $employee);
		} else {
			$stats->{$dienst}->{'certified'}->{count} += 1;
			push (@{$stats->{$dienst}->{'certified'}->{list}}, $employee);
		}
	}
	
	
	#print Dumper($stats);
	
	# Display the results
	foreach my $dienst (sort keys $stats) {
		next if ($dienst eq 'employee_count');
		
		if (!defined $stats->{$dienst}->{certified}->{count}) { $stats->{$dienst}->{certified}->{count} = 0};
		if (!defined $stats->{$dienst}->{training}->{count}) { $stats->{$dienst}->{training}->{count} = 0};
		if (!defined $stats->{$dienst}->{not_started}->{count}) { $stats->{$dienst}->{not_started}->{count} = 0};
		
		say "$dienst \t: " . $stats->{$dienst}->{certified}->{count} ." -- " . $stats->{$dienst}->{training}->{count} . " -- " . $stats->{$dienst}->{not_started}->{count};
		
	}
}

# Parse the employee database
sub _employee_parser {
	my ($self, $fname) = @_;
	
	my $employees;
	
	my $converter = Text::Iconv -> new ("utf-8", "windows-1251");
	my $excel = Spreadsheet::XLSX -> new ($self->{employees}, $converter);	
	
	foreach my $sheet (@{$excel -> {Worksheet}}) {
 
       #printf("Sheet: %s\n", $sheet->{Name});
        
       $sheet -> {MaxRow} ||= $sheet -> {MinRow};
        
        # Go over the rows in the sheet and extract employee info, skip first row
        foreach my $row ($sheet -> {MinRow} + 1 .. $sheet -> {MaxRow}) {
         
         	my $dienst = $sheet->{Cells}[$row][0]->{Val};
         	my $familyname = uc($sheet->{Cells}[$row][2]->{Val});
         	my $givenname  = uc($sheet->{Cells}[$row][3]->{Val});
 
 			my $name = "$familyname $givenname";
 			$employees->{$name} = {dienst => $dienst};
 			
		}
 
    }
	return $employees;
	
}
# Speed up the Moose object construction
__PACKAGE__->meta->make_immutable;
no Moose;
1;

# ABSTRACT: add description

=head1 SYNOPSIS

my $object = App::CPRReporter->new(parameter => 'text.txt');

=head1 DESCRIPTION

Describe the module here

=head1 METHODS

=head2 C<new(%parameters)>

This constructor returns a new App::CPRReporter object. Supported parameters are listed below

=over

=item parameters

Describe

=back

=head2 BUILD

Helper function to run custome code after the object has been created by Moose.

=cut

