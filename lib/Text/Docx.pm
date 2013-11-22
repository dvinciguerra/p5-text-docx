package Text::Docx;

use strict;
use warnings;

use File::Temp;
use Archive::Zip;

# module version
our $VERSION = '0.20';

# constructor
sub new {
    my ($class, $filename) = @_;

    # new temp file
    my $tmp = $class->_new_tempfile;
    my $fname = $tmp->filename;

    # template to temp
    open my $fh, "<", "$filename" or die "error: $!";
    print $tmp $_ while <$fh>;
    close $fh;

    $tmp->seek( 0, SEEK_END );

    # open docx
    my $zip = $class->_open_docx($fname);

    return bless {
        _docx => $zip || undef,
        _temp_filename => $fname || undef
    }, $class;
}

# accessors
sub temp_filename {
    return shift->{_temp_filename};
}

# methods
sub process {
    my ($self, $stash) = @_;

    # error
    die "error: docx archive handler is closed" 
        unless $self->{_docx};

    # getting docx file
    my $zip = $self->{_docx};
    my $content = $zip->contents( 'word/document.xml' );

    # replace stash vars
    $content =~ s/\[%\s*$_\s*%\]/$stash->{$_}/g for keys %$stash;

    # update docx file
    $zip->contents( 'word/document.xml', $content );
    $zip->writeToFileNamed($self->temp_filename);

    $self;
}

sub save_to {
    my ($self, $filename) = @_;

    $self->{_docx}->writeToFileNamed($filename) 
        or die "error: $!";

    $self;
}

# private methods
sub _open_docx {
    my ($self, $filename) = @_;
    return Archive::Zip->new($filename)
        or die "error: $!";
}

sub _new_tempfile {
    my $self = shift;
    return File::Temp->new(UNLINK => 0);
}

1;
__END__

=pod

=head1 NAME

Text::Docx - Docx files as Template

=head1 DESCRIPTION


=head1 SYNOPSIS

    use Text::Docx;

    # getting a new instance
    my $template = Text::Docx->new( '/path/to/template.docx' );

    # basic stash
    my $stash = {
        foo => 'foo value', bar => 'bar value',
    };

    # process template and save docx file
    $template->process( $stash )
        ->save_to( '/path/to/new/file.docx' );

and in docx file we have something like:

    You is the [% foo %] of my [% bar %]!!!


=head2 Methods


=head1 BUGS

=head1 AUTHOR

2013 (c) Bivee L<http://bivee.com.br>

Daniel Vinciguerra <daniel.vinciguerra@bivee.com.br>


=head1 COPYRIGHT AND LICENSE

This software is copyright (c) 2013 by Bivee.

This is a free software; you can redistribute it and/or modify it under 
the same terms of Perl 5 programming languagem system itself.

=cut
