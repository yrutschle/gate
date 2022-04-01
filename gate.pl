#! /usr/bin/perl -w

# Copyright © 2017-2022 Yves Rűtschlé
# 
# Permission is hereby granted, free of charge, to any person obtaining a copy of
# this software and associated documentation files (the “Software”), to deal in
# the Software without restriction, including without limitation the rights to
# use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies
# of the Software, and to permit persons to whom the Software is furnished to do
# so, subject to the following conditions:
# 
# The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.
# 
# THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.

use Excel;
use Getopt::Long;
use Pod::Usage;

=head1 NAME

 gate -- Generic Array Translation and exploitation

=head1 SYNOPSIS

 gate [--data <data worksheet name>] [--style <stype worksheet name>] tab_file.xls
 gate --help

=head1 DESCRIPTION

B<gate> translates data from an Excel spreadsheet into any other format, as
long as the output data can be expressed using simple substitutions and stored
into an Excel cell.

Please keep in mind that a 'worksheet' is the official Excel name for a tab.
Also note that Excel probably needs to be running for this program to work
properly.

=head2 Data worksheet

The first line of the data sheet must contain the name of parameters contained
in each column, and the first column must contain a 'type' parameter. I<gate>
then reads each line of the data sheet, and outputs the corresponding model
from the style worksheet.

=head2 Style worksheet

Each line of the Style worksheet contains one model. The first cell contains
the name of the model (which corresponds to the 'type' specified in the data
sheet), and the second cell contains the model proper. Parameters can be
specified with a dollar sign. Parameter names should correspond to the names of
the first line of the data sheet. Note that parameter names can also refer to
another model, which allows to build complex models. Using recursive models is
not advised.

Special model names B<pre_out> and B<post_out> can be defined to control text
that is output before and after any data line is processed. This can be used to
create, for example, an entire, valid HTML page. If not defined, these default
to create an HTML and BODY envelope.

B<gate> also has built-in types B<h1> to B<h6> which create HTML titles, and
B<text> type which simply copies data through. These can be redefined in the
style worksheet if the defaults are not wanted.

=head1 OPTIONS

=over 4

=item B<--data data worksheet name>

Name of the data worksheet. Defaults to 'Data'.

=item B<--style [excel file:]mode_tab_name[:tab2[:tab3...]]>

Name of the style worksheet. Defaults to 'Styles'.
Warning: if the Excel file is not in the current directory, then its path must
be absolute. Also, if you specify an Excel filename, it must be different than
the data file.

Several tabs can be specified in a row, in which case styles are read from each
tab in turn. This allows to create an Excel file containing generic styles in
one tab, then picking one or several additional tabs, either to override the
generic styles or to add new ones.

=back

=head1 EXAMPLES

 gate.pl doc.xlsx > out.html

Processes data contained in tab 'Data' through styles contained in 'Styles',
save output in out.html.

 gate.pl doc.xlsx --data MyData --styles styles.xlsx;MyStyles;NewStyles > out.html

Processes data contained in tab 'MyData' of doc.xlsx through styles contained in
'MyStyles' tab, overriden by 'NewStyles' of 'styles.xlsx' file, save output in out.html.

=cut

my ($debug) = 1;

my ($param_data_tab, $param_style, $help);
GetOptions(
    'data=s' => \$param_data_tab,
    'style=s' => \$param_style,
    'help' => \$help,
) or die pod2usage();


my $excel_filename = shift;
$param_style //= "Styles";
$param_data_tab //= "Data";

die pod2usage(-verbose=>2) if $help;
die pod2usage() unless defined $excel_filename;


my %template;  # associates template name to corresponding HTML; this is read from the "Styles" tab of the Excel sheet

my %vars; # associates keywords ('variables') with their output

my $excel = Excel->new($excel_filename);

# Read the templates
my $style_filename;
my @style_tabs = split /;/, $param_style;
$style_filename = shift @style_tabs if $style_tabs[0] =~ /\.xls/;

my $excel_style = $excel;
warn "Sourcing styles from $style_filename\n" if $debug;
$excel_style = Excel->new($style_filename) if defined $style_filename;
foreach my $param_style_tab (@style_tabs) {
    warn "Loading styles from tab $param_style_tab\n" if $debug;
    foreach my $columns ($excel_style->sheet($param_style_tab)) {
        my ($type, $template) = @$columns;
        if (defined $type) {
            $template //= '';
            $template{$type} = $template; 
            $vars{$type} = $template;
            warn "Defined '$type': $template\n" if $debug;
        }
    }
}

sub quotehtml {
    my($out) = @_;
    return "" unless defined $out;
    $out =~ s/\n/<br>/gs;
    return $out;
}

# Replace variables in data, recursively
sub recursive_interpolate {
    my ($vars, $data) = @_;
    return undef if not defined $data;

    $vars //= {};

    while ($data =~ /\$(\w+)/) {
        my $varname = $1;
        if (not defined $vars->{$varname}) {
            warn "\$$varname does not exist\n";
            $vars->{$varname} = "";
        }
        $data =~ s/\$$varname/$vars->{$varname}/;
    }
    return $data;
}

my @lines = $excel->sheet($param_data_tab);

# The first line of 'Tests' document the variable names
my @vars = @{shift @lines};
shift @vars;  # and the first column contains 'Type', which we don't care about

print recursive_interpolate(undef, $template{pre_out}) // "<html><body>";

my $cnt = 0;
foreach my $columns (@lines) {
    $cnt++;
    my ($type) = shift @$columns;
    next if not defined $type;
    warn "line $cnt: $type\n" if $debug;
    if (exists $template{$type}) {
        foreach my $var (@vars) {
            $var //= "undef"; # If the cell is empty, the value is undef;
                              # We still need to shift to the next column,
                              # but don't want to get a warning for undef $var
            $vars{$var} = quotehtml shift @$columns;
        }

        # Start from the template and substitute all variables
        my $out = $template{$type};
        $out = recursive_interpolate(\%vars, $out);
        print $out; 
    } elsif ($type =~ /^h(\d)/) {
        my ($title) = @$columns;
        print "<h$1>$title</h$1>\n";
    } elsif ($type eq 'text') {
        my ($text) = @$columns;
        $text =~ s/\n/<br>/gs;
        print "<p class=Text>$text</p>\n";
    } else {
        warn "type '$type' not defined\n";
    }
}


print recursive_interpolate(undef, $template{post_out}) // "</body></html>\n";
