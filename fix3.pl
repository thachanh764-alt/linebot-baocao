#!/usr/bin/perl
use strict;
use warnings;

local $/;
open(my $fh, '<', $ARGV[0]) or die $!;
my $content = <$fh>;
close($fh);

$content =~ s/(console\.log\('\[DEBUG2\][^\n]*\);)/$1\n  const badChars = [];\n  for (let i = 0; i < token.length; i++) {\n    const code = token.charCodeAt(i);\n    if (code < 32 || code > 126) badChars.push({ pos: i, code });\n  }\n  console.log('[DEBUG3] ky tu la trong token:', JSON.stringify(badChars));/;

open(my $out, '>', $ARGV[0]) or die $!;
print $out $content;
close($out);
