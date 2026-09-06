#!/usr/bin/perl
use strict;
use warnings;

local $/;
open(my $fh, '<', $ARGV[0]) or die $!;
my $content = <$fh>;
close($fh);

$content =~ s/(const token = \(process\.env\.LINE_CHANNEL_ACCESS_TOKEN \|\| ''\)\.trim\(\);)/$1\n  console.log('[DEBUG2] messageId:', messageId, '- token length:', token.length, '- 6 ky tu cuoi:', token.slice(-6));/;

open(my $out, '>', $ARGV[0]) or die $!;
print $out $content;
close($out);
