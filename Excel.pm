package Excel;


# from http://docs.activestate.com/activeperl/5.10/faq/Windows/ActivePerl-Winfaq12.html
use Win32::OLE qw(in with);
use Win32::OLE::Const 'Microsoft Excel';
$Win32::OLE::Warn = 3;                                # die on errors...

@ISA = qw/Win32::OLE/;

=head1 @array = read("foo.xls", $worksheet);

Reads the active zone in the specified worksheet of the  spreadsheet.
Filename should be either absolute or in cwd.

=cut
use Data::Dumper;
sub read {
    my ($class, $filename, $worksheet) = @_;

    my $Excel = Win32::OLE->GetActiveObject('Excel.Application')
    || Win32::OLE->new('Excel.Application', 'Quit');

    unless ($filename =~ /\\/) {
        my $wd = `cd`;
        chomp $wd;
        $filename = "$wd\\$filename";
    }
    my $Book = $Excel->Workbooks->Open($filename);
    my $Sheet = $Book->Worksheets($worksheet);

    my $last_cell = $Sheet->Range("A1")->EntireColumn->
        SpecialCells(xlCellTypeLastCell)->{Address};
    $last_cell =~ s/\$//g;
    my $range = "A1:$last_cell";

    my $array = $Sheet->Range($range)->{'Value'};
    $Book->Close;
    return @$array;
}

# Excel->read is legacy, it can only be called once which isn't so good.
# If you need to read several tabs from an Excel, use the following functions


# Returns an Excel object of sorts
sub new {
    my ($class, $filename) = @_;

    my $Excel = Win32::OLE->GetActiveObject('Excel.Application')
    || Win32::OLE->new('Excel.Application', 'Quit');

    unless ($filename =~ /\\/) {
        my $wd = `cd`;
        chomp $wd;
        $filename = "$wd\\$filename";
    }
    my $Book = $Excel->Workbooks->Open($filename);

    return bless $Book, $class;
}

# Returns the lines in a sheet as one big array
sub sheet {
    my ($Book, $worksheet) = @_;

    my $Sheet = $Book->Worksheets($worksheet);

    die "Error loading the workbook $Book. Did you start Excel?\n" unless defined $Sheet;

    my $last_cell = $Sheet->Range("A1")->EntireColumn->
        SpecialCells(xlCellTypeLastCell)->{Address};
    $last_cell =~ s/\$//g;
    my $range = "A1:$last_cell";

    my $array = $Sheet->Range($range)->{'Value'};
    return @$array;
}
