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
use Text::Fuzzy::PP;

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
    # Make an array of employees that will be used for fuzzy matching
    foreach my $employee (keys $self->{_employees}) {
	push (@{$self->{_employee_array}}, $employee);
    }
    print Dumper($self->{_employee_array});
}

# Run the application, merging the info of the certificates and the employees
sub run {
	my $self = shift;
	
	# Certificates are here
	# TODO adapt the parser module so that the level of hierarchy in this hash is less deep
	my $certificate_count = 0;
	my $certs = $self->{_certificates}->{_data}->{certs};
	foreach my $date (sort keys $certs) {
		foreach my $certuser (@{$certs->{$date}}){
			my $fullname = $self->_resolve_name($certuser->{familyname}, $certuser->{givenname});
			#say "Certificate found for $fullname";
			$certificate_count++;

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

	say "$certificate_count certificates found";
	
	my $training_count = 0;
	my $training = $self->{_certificates}->{_data}->{training};
	foreach my $traininguser (@{$training}) {
		my $fullname = $self->_resolve_name($traininguser->{familyname}, $traininguser->{givenname});
		#say "Training found for $fullname";
			# TODO deduplicate this code with a local function, see above
			if (defined $self->{_employees}->{$fullname}){
				# Fill in training if there is no certificate yet, otherwise notify!
				if (!defined $self->{_employees}->{$fullname}->{cert}){
					$self->{_employees}->{$fullname}->{cert} = 'training';
					$training_count++;
				} else {
				#	carp "Warning: employee '$fullname' is both in training and has a certificate from $self->{_employees}->{$fullname}->{cert}";
				}
			} else {
				# Oops: user not found in personel database
				carp "Warning: employee '$fullname' not found in employee database";
				$training_count++;
			}
	}
	
	say "$training_count people are in training";

	# Check people who are in training and that have a certificate
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
		
		say "$dienst;" . $stats->{$dienst}->{certified}->{count} .";" . $stats->{$dienst}->{training}->{count} . ";" . $stats->{$dienst}->{not_started}->{count};
		
	}

	print Dumper($self->{_resolve});
	#print Dumper($stats);
}

# Parse the employee database
sub _employee_parser {
	my ($self, $fname) = @_;
	
	my $employees;
	
	#my $converter = Text::Iconv -> new ("utf-8", "windows-1251");
	my $excel = Spreadsheet::XLSX -> new ($self->{employees});	
	
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

# Try to resolve a name in case it is not found in the personel database
sub _resolve_name {
	my ($self, $fname, $gname) = @_;

	my $name;

	# Cleanup leading/trailing spaces
	# Todo move this to module, data we receive here should be clean
	$fname =~ s/^\s+//; # strip white space from the beginning
        $fname =~ s/\s+$//; # strip white space from the end
	$gname =~ s/^\s+//; # strip white space from the beginning
        $gname =~ s/\s+$//; # strip white space from the end

	# Straight match
	$name = uc($fname) . " " . uc($gname);
	if (exists $self->{_employees}->{$name}) {
		$self->{_resolve}->{straight} += 1;
		return $name;
	}

	# First try, maybe they switched familyname and givenname?
	$name = uc($gname) . " " . uc($fname);
	if (exists $self->{_employees}->{$name}) {
		my $fixed = $self->{_employees}->{$name};
		$self->_fixlog('switcharoo', $name, $fixed) ;
		return $fixed; 
	}

	# Exact match but missing parts?
	$name = uc($fname) . " " . uc($gname);
	foreach my $employee (@{$self->{_employee_array}}) {
		if ($employee =~ /.*$name.*/) {
			$self->_fixlog('partial', $name, $employee);
			return $employee;
		}
	}

	# Check if we can find a match with fuzzy matching
	$name = uc($fname) . " " . uc($gname);
	my $tf = Text::Fuzzy::PP->new ($name);
	$tf->set_max_distance(3);
	my $index = $tf->nearest ($self->{_employee_array}) || -1;
	if ($index > 0) {
		my $fixed = $self->{_employee_array}->[$index]; 
		$self->_fixlog('fuzzy', $name, $fixed);
		return $fixed;
	}
	
	# Report no match found
	#say "No match in employee database for '$name'";
	$self->{_resolve}->{nomatch} += 1;
	return $name;
		
}

# Log resolved names so that they can be used for later reference
sub _fixlog {
	my ($self, $type, $original, $fixed) = @_;
	
	#say "$type match for '$original', replaced by '$fixed'";
	$self->{_resolve}->{$type} += 1;
	push (@{$self->{_resolve_list}}, {$original => {fixed => $fixed, type => $type}});

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

