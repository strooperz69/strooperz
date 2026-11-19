$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$outPath = Join-Path (Split-Path $PSScriptRoot -Parent) 'nist-simple\GCC-High-CONMON-Wave1-Playbook.docx'
$tmp = Join-Path $env:TEMP ('conmon-docx-' + [guid]::NewGuid().ToString('n'))
New-Item -ItemType Directory -Path (Join-Path $tmp 'word\_rels') | Out-Null
New-Item -ItemType Directory -Path (Join-Path $tmp '_rels') | Out-Null

function Esc([string]$s) {
    if ($null -eq $s) { return '' }
    ($s -replace '&', '&amp;' -replace '<', '&lt;' -replace '>', '&gt;')
}

$sb = New-Object System.Text.StringBuilder
[void]$sb.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
[void]$sb.Append('<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>')

function Pxml([string]$text, [string]$style) {
    $t = Esc $text
    if ($style) {
        return "<w:p><w:pPr><w:pStyle w:val=`"$style`"/></w:pPr><w:r><w:t xml:space=`"preserve`">$t</w:t></w:r></w:p>"
    }
    return "<w:p><w:r><w:t xml:space=`"preserve`">$t</w:t></w:r></w:p>"
}
function QueryXml([string]$code) {
    $out = New-Object System.Text.StringBuilder
    foreach ($line in ($code -split "`r?`n")) {
        $t = Esc $line
        if ($t -eq '') { $t = ' ' }
        [void]$out.Append("<w:p><w:pPr><w:pStyle w:val=`"Query`"/><w:shd w:val=`"clear`" w:fill=`"F2F2F2`"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii=`"Consolas`" w:hAnsi=`"Consolas`"/><w:sz w:val=`"18`"/></w:rPr><w:t xml:space=`"preserve`">$t</w:t></w:r></w:p>")
    }
    return $out.ToString()
}

$blocks = @(
    @{ k = 'h1'; v = 'GCC High CONMON  -  Wave 1 playbook' }
    @{ k = 'p'; v = 'For: Cybersecurity Engineer. Time: a few hours to two days. Tool: SentinelOne Event Search (PowerQuery).' }
    @{ k = 'p'; v = 'Job: prove we can see Global Administrator (GA) and Privileged Role Administrator (PRA) changes in SentinelOne, then save two queries for weekly review. That is the whole assignment. Stop when the checklist at the bottom is done.' }
    @{ k = 'p'; v = 'Do not treat a Global Admin change as an incident. Match it to a change ticket. If there is no ticket, tell the CONMON owner. Do not open IR on your own.' }

    @{ k = 'h2'; v = 'Before you start' }
    @{ k = 'b'; v = 'Work only in SentinelOne Event Search. Do not start in Purview or Log Analytics.' }
    @{ k = 'b'; v = 'Set the time picker in the UI (try Last 7 days). Do not put a time filter in the query.' }
    @{ k = 'b'; v = 'Every query must start with: tag != ''logVolume''' }
    @{ k = 'b'; v = 'Do not use | columns *' }
    @{ k = 'b'; v = 'Do not put = * inside a let line' }
    @{ k = 'b'; v = 'Copilot usage monitoring is already built. Do not work on Copilot in this wave.' }

    @{ k = 'h1'; v = 'Step 1  -  Find the log source' }
    @{ k = 'p'; v = 'Paste this. Run it. You should get a small table of dataSource.name and a count.' }
    @{ k = 'q'; v = @"
tag != 'logVolume'
dataSource.name = *
(
  * contains 'Global Administrator'
  OR * contains 'Privileged Role Administrator'
  OR * contains 'Add member to role'
  OR * contains 'Add eligible member to role'
  OR * contains 'Privileged Identity Management'
)
| group n=count() by dataSource.name
| sort -n
"@ }
    @{ k = 'p'; v = 'Write down the source names and counts.' }
    @{ k = 'p'; v = 'How to pick the source for the rest of this playbook:' }
    @{ k = 'b'; v = 'Prefer a source whose name is Microsoft Entra ID (or similar Entra / Azure AD audit). That is identity.' }
    @{ k = 'b'; v = 'If you also see Azure Event Hubs, do not use it for the GA list. Open one event. It is often PIM approvals or hub noise. Leave it. Come back later if asked.' }
    @{ k = 'b'; v = 'Ignore a row where dataSource.name is blank (null). That is S1 recording your own search, not tenant logs.' }
    @{ k = 'p'; v = 'If you get zero named sources: stop. Tell the CONMON owner. The next check is whether Entra AuditLogs are on the Event Hub S1 reads. Do not guess. Do not switch to Purview as the monitoring tool.' }
    @{ k = 'p'; v = 'From here on, SOURCE means the Entra name you wrote down. In the queries below it is shown as Microsoft Entra ID. If your name is different, replace it.' }

    @{ k = 'h1'; v = 'Step 2  -  List the changes' }
    @{ k = 'p'; v = 'Paste this. Same time picker.' }
    @{ k = 'q'; v = @"
tag != 'logVolume'
dataSource.name = 'Microsoft Entra ID'
(
  * contains 'Global Administrator'
  OR * contains 'Privileged Role Administrator'
)
| let Actor = ActorName ? ActorName : (UserId ? UserId : unmapped.initiatedBy)
| let Op = Operation ? Operation : (activityDisplayName ? activityDisplayName : unmapped.activityDisplayName)
| filter !isempty(Op)
| columns timestamp, Actor, Op
| sort -timestamp
"@ }
    @{ k = 'p'; v = 'Success: you get a table of time, person, and action.' }
    @{ k = 'p'; v = 'If Actor or Op is blank, do not rewrite the query from scratch. Open one raw event. Find the person field and the activity field. Put those exact names in the two let lines. Then run again.' }
    @{ k = 'p'; v = 'If PowerQuery is 0 but raw events exist, the filter dropped the rows. Run this, open 5 events, then put the real field names into Step 2:' }
    @{ k = 'q'; v = @"
tag != 'logVolume'
dataSource.name = 'Microsoft Entra ID'
(
  * contains 'Global Administrator'
  OR * contains 'Privileged Role Administrator'
)
| limit 5
"@ }
    @{ k = 'p'; v = 'A large raw count with few useful Ops is normal. The word search hits extra JSON. Step 3 will show which operations are real.' }
    @{ k = 'p'; v = 'Save this search as: CONMON W1-02 list' }

    @{ k = 'h1'; v = 'Step 3  -  Count by person and action' }
    @{ k = 'p'; v = 'Paste this. Same source name. Same time picker.' }
    @{ k = 'q'; v = @"
tag != 'logVolume'
dataSource.name = 'Microsoft Entra ID'
(
  * contains 'Global Administrator'
  OR * contains 'Privileged Role Administrator'
)
| let Actor = ActorName ? ActorName : (UserId ? UserId : unmapped.initiatedBy)
| let Op = Operation ? Operation : (activityDisplayName ? activityDisplayName : unmapped.activityDisplayName)
| filter !isempty(Op)
| group n=count() by Actor, Op
| sort -n
"@ }
    @{ k = 'p'; v = 'This is the weekly review table. Give it to GRC with the change-ticket list. They match names. You do not need to explain CMMC.' }
    @{ k = 'p'; v = 'If almost every row looks like noise, tell the CONMON owner and send a screenshot of the top Op values. Do not invent a tighter filter on your own yet.' }
    @{ k = 'p'; v = 'Save this search as: CONMON W1-03 by actor' }

    @{ k = 'h1'; v = 'Step 4  -  Who is allowed to be GA or PRA' }
    @{ k = 'p'; v = 'This is not a SentinelOne query. In Entra, open Privileged Identity Management (or Roles). For Global Administrator and Privileged Role Administrator, make a small spreadsheet:' }
    @{ k = 'b'; v = 'UPN' }
    @{ k = 'b'; v = 'Eligible or standing Active' }
    @{ k = 'b'; v = 'Cloud-only account or synced from on-prem' }
    @{ k = 'b'; v = 'Break-glass account? yes or no' }
    @{ k = 'p'; v = 'What we want: people are Eligible (they must activate), not standing Active. Break-glass is documented and cloud-only. If you find standing Active humans who are not break-glass, write it down. Do not remove access yourself.' }

    @{ k = 'h1'; v = 'Step 5  -  One sentence in the runbook' }
    @{ k = 'p'; v = 'Add this exact sentence to the SOC or GRC runbook:' }
    @{ k = 'q'; v = 'Global Administrator and Privileged Role Administrator activity in SentinelOne is reviewed against change tickets. Presence of the event is not an incident.' }

    @{ k = 'h1'; v = 'If logs are missing' }
    @{ k = 'p'; v = 'Only if Step 1 found no named source, or Step 2 cannot show any GA/PRA activity in a week when you know work happened:' }
    @{ k = 'n'; v = 'Ask IT: are Entra AuditLogs sent to the Event Hub that SentinelOne reads? Copilot logs on that hub do not mean Entra is on it.' }
    @{ k = 'n'; v = 'If the hub has the events but S1 cannot query fields, it is a parser gap. Keep using the queries above. File a parser ticket with one sample event. Do not wait on the parser to finish Wave 1.' }
    @{ k = 'n'; v = 'Log Analytics is only a yes/no: table AuditLogs, did Microsoft emit the event. If yes in Log Analytics and no in S1, ingest is broken. If no in both, Entra is not logging to the hub.' }
    @{ k = 'n'; v = 'Do not build this program in Purview.' }

    @{ k = 'h1'; v = 'Done. Stop here.' }
    @{ k = 'b'; v = 'Step 1 source name written down. Entra source used. Event Hubs source not used for the GA list.' }
    @{ k = 'b'; v = 'CONMON W1-02 list saved and returns time, person, action.' }
    @{ k = 'b'; v = 'CONMON W1-03 by actor saved.' }
    @{ k = 'b'; v = 'Spreadsheet of GA and PRA assignments complete.' }
    @{ k = 'b'; v = 'Runbook sentence added.' }
    @{ k = 'p'; v = 'Hand the two saved queries and the spreadsheet back. Do not add alerts, Copilot, Exchange, or Netskope until you are told Wave 2 has started.' }

    @{ k = 'h1'; v = 'Wave 2 (not your job yet)' }
    @{ k = 'p'; v = 'Later we will add eight SentinelOne alerts: GA activation, PRA assign/activate, break-glass sign-in, standing active GA, log pipeline change, silent source, Copilot/app consent, Conditional Access change. Do not create them in Wave 1.' }
)

foreach ($item in $blocks) {
    switch ($item.k) {
        'h1' { [void]$sb.Append((Pxml $item.v 'Heading1')) }
        'h2' { [void]$sb.Append((Pxml $item.v 'Heading2')) }
        'p'  { [void]$sb.Append((Pxml $item.v $null)) }
        'b'  { [void]$sb.Append((Pxml $item.v 'ListBullet')) }
        'n'  {
            $t = Esc ($item.v)
            [void]$sb.Append("<w:p><w:pPr><w:pStyle w:val=`"ListNumber`"/><w:numPr><w:ilvl w:val=`"0`"/><w:numId w:val=`"1`"/></w:numPr></w:pPr><w:r><w:t xml:space=`"preserve`">$t</w:t></w:r></w:p>")
        }
        'q'  { [void]$sb.Append((QueryXml $item.v)) }
    }
}

[void]$sb.Append('<w:sectPr><w:pgMar w:top="1080" w:right="1080" w:bottom="1080" w:left="1080"/></w:sectPr>')
[void]$sb.Append('</w:body></w:document>')
[System.IO.File]::WriteAllText((Join-Path $tmp 'word\document.xml'), $sb.ToString(), [System.Text.UTF8Encoding]::new($true))

$styles = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr><w:pPr><w:spacing w:after="160"/></w:pPr></w:style>
  <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:before="360" w:after="160"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:before="240" w:after="120"/><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/><w:sz w:val="26"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="ListBullet"><w:name w:val="List Bullet"/><w:basedOn w:val="Normal"/><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr></w:pPr></w:style>
  <w:style w:type="paragraph" w:styleId="ListNumber"><w:name w:val="List Number"/><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="Query"><w:name w:val="Query"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="0" w:after="0"/><w:ind w:left="144"/></w:pPr></w:style>
</w:styles>
'@
[System.IO.File]::WriteAllText((Join-Path $tmp 'word\styles.xml'), $styles, [System.Text.UTF8Encoding]::new($true))

$numbering = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="1">
    <w:multiLevelType w:val="singleLevel"/>
    <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>
  </w:abstractNum>
  <w:abstractNum w:abstractNumId="2">
    <w:multiLevelType w:val="singleLevel"/>
    <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="-"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num>
  <w:num w:numId="2"><w:abstractNumId w:val="2"/></w:num>
</w:numbering>
'@
[System.IO.File]::WriteAllText((Join-Path $tmp 'word\numbering.xml'), $numbering, [System.Text.UTF8Encoding]::new($true))

$rels = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
'@
[System.IO.File]::WriteAllText((Join-Path $tmp '_rels\.rels'), $rels, [System.Text.UTF8Encoding]::new($true))

$docRels = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
</Relationships>
'@
[System.IO.File]::WriteAllText((Join-Path $tmp 'word\_rels\document.xml.rels'), $docRels, [System.Text.UTF8Encoding]::new($true))

$ct = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
</Types>
'@
[System.IO.File]::WriteAllText((Join-Path $tmp '[Content_Types].xml'), $ct, [System.Text.UTF8Encoding]::new($true))

if (Test-Path $outPath) { Remove-Item $outPath -Force }
[System.IO.Compression.ZipFile]::CreateFromDirectory($tmp, $outPath, [System.IO.Compression.CompressionLevel]::Optimal, $false)
Remove-Item $tmp -Recurse -Force
$dl = Join-Path $env:USERPROFILE 'Downloads'
if (Test-Path $dl) {
    Copy-Item $outPath (Join-Path $dl 'GCC-High-CONMON-Wave1-Playbook.docx') -Force
}
Get-Item $outPath | Select-Object FullName, Length, LastWriteTime
