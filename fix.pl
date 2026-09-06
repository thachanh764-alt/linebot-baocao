#!/usr/bin/perl
use strict;
use warnings;

local $/;
open(my $fh, '<', $ARGV[0]) or die $!;
my $content = <$fh>;
close($fh);

my $new_func = <<'NEWFUNC';
async function taiNoiDungFileLine(messageId) {
  const token = (process.env.LINE_CHANNEL_ACCESS_TOKEN || '').trim();
  const res = await fetch(`https://api-data.line.me/v2/bot/message/${messageId}/content`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!res.ok) {
    const body = await res.text();
    throw new Error(`Tai file LINE that bai: ${res.status} ${body}`);
  }
  const arrayBuffer = await res.arrayBuffer();
  return Buffer.from(arrayBuffer);
}
NEWFUNC
chomp $new_func;

$content =~ s/async function taiNoiDungFileLine\(messageId\) \{.*?\n\}/$new_func/s;

open(my $out, '>', $ARGV[0]) or die $!;
print $out $content;
close($out);
