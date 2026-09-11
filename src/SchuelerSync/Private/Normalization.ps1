function ConvertTo-UpnToken {
    param([Parameter(Mandatory)][string] $Value)
    $mapped = $Value.Trim().ToLowerInvariant().
        Replace('ä', 'ae').Replace('ö', 'oe').Replace('ü', 'ue').Replace('ß', 'ss')
    $decomposed = $mapped.Normalize([Text.NormalizationForm]::FormD)
    $builder = [Text.StringBuilder]::new()
    foreach ($character in $decomposed.ToCharArray()) {
        if ([Globalization.CharUnicodeInfo]::GetUnicodeCategory($character) -ne
            [Globalization.UnicodeCategory]::NonSpacingMark) {
            [void] $builder.Append($character)
        }
    }
    return ($builder.ToString().Normalize([Text.NormalizationForm]::FormC) -replace '[^a-z0-9]', '')
}

function Get-OfficeLocation {
    param([Parameter(Mandatory)][string] $ClassName)
    $matches = [regex]::Matches($ClassName, '(?i)(?<office>g[1-4]|m[1-3]|o[1-2]|a[1-2])(?=_|$)')
    if ($matches.Count -ne 1) {
        throw "Klasse '$ClassName' enthält keine eindeutige erlaubte Office Location."
    }
    return $matches[0].Groups['office'].Value.ToUpperInvariant()
}

function Get-UpnCandidates {
    param(
        [Parameter(Mandatory)][string] $GivenName,
        [Parameter(Mandatory)][string] $Surname,
        [Parameter(Mandatory)][string] $Domain
    )
    $givenToken = ConvertTo-UpnToken $GivenName
    $surnameToken = ConvertTo-UpnToken $Surname
    if ([string]::IsNullOrEmpty($givenToken) -or [string]::IsNullOrEmpty($surnameToken)) {
        throw 'Vor- und Nachname müssen nach der Normalisierung nicht leer sein.'
    }
    for ($length = 1; $length -le $givenToken.Length; $length++) {
        '{0}{1}@{2}' -f $givenToken.Substring(0, $length), $surnameToken, $Domain.Trim()
    }
}

function Get-AddressOwnerIds {
    param([AllowNull()][object] $Value)
    if ($null -eq $Value) { return @() }
    if ($Value -is [string] -or $Value -is [Guid]) { return @([string]$Value) }
    if ($Value -is [System.Collections.IDictionary]) {
        return @($Value.Values | ForEach-Object { Get-AddressOwnerIds $_ })
    }
    if ($Value -is [System.Collections.IEnumerable] -and $Value -isnot [hashtable]) {
        return @($Value | ForEach-Object { Get-AddressOwnerIds $_ })
    }
    $ids = @()
    foreach ($name in 'CurrentObjectId', 'GraphObjectId', 'ExchangeObjectId', 'ObjectId', 'OwnerId', 'OwnerIds', 'Ids') {
        $property = $Value.PSObject.Properties[$name]
        if ($null -ne $property -and $null -ne $property.Value) {
            $ids += @(Get-AddressOwnerIds $property.Value)
        }
    }
    return @($ids | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
}

function Test-UsedAddress {
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][System.Collections.Generic.HashSet[string]] $UsedAddresses,
        [Parameter(Mandatory)][string] $Candidate
    )
    foreach ($address in $UsedAddresses) {
        if ([string]::Equals([string]$address, $Candidate, [StringComparison]::OrdinalIgnoreCase)) {
            return $true
        }
    }
    return $false
}

function Select-AvailableUpn {
    param(
        [Parameter(Mandatory)][string] $GivenName,
        [Parameter(Mandatory)][string] $Surname,
        [Parameter(Mandatory)][string] $Domain,
        [Parameter(Mandatory)][AllowEmptyCollection()][System.Collections.Generic.HashSet[string]] $UsedAddresses,
        [Parameter(Mandatory)][AllowEmptyCollection()][System.Collections.IDictionary] $AddressOwners,
        [AllowNull()][string] $CurrentObjectId
    )
    $candidates = @(Get-UpnCandidates $GivenName $Surname $Domain)
    $collisions = [Collections.Generic.List[string]]::new()
    $isAvailable = {
        param($candidate)
        if (Test-UsedAddress -UsedAddresses $UsedAddresses -Candidate $candidate) { return $false }
        $ownerEntry = $null
        foreach ($key in $AddressOwners.Keys) {
            if ([string]::Equals([string]$key, $candidate, [StringComparison]::OrdinalIgnoreCase)) {
                $ownerEntry = $AddressOwners[$key]
                break
            }
        }
        if ($null -eq $ownerEntry) { return $true }
        $owners = @(Get-AddressOwnerIds $ownerEntry)
        if ([string]::IsNullOrWhiteSpace($CurrentObjectId) -or $owners.Count -eq 0) { return $false }
        foreach ($owner in $owners) {
            if (-not [string]::Equals($owner, $CurrentObjectId, [StringComparison]::OrdinalIgnoreCase)) { return $false }
        }
        return $true
    }
    foreach ($candidate in $candidates) {
        if (& $isAvailable $candidate) {
            [void] $UsedAddresses.Add($candidate)
            return [pscustomobject]@{ Upn = $candidate; WasFallback = $false; CollisionCount = $collisions.Count; Collisions = [string[]]$collisions }
        }
        $collisions.Add($candidate)
    }
    $fullGiven = ConvertTo-UpnToken $GivenName
    $surnameToken = ConvertTo-UpnToken $Surname
    for ($suffix = 2; $true; $suffix++) {
        $candidate = '{0}{1}{2}@{3}' -f $fullGiven, $surnameToken, $suffix, $Domain.Trim()
        if (& $isAvailable $candidate) {
            [void] $UsedAddresses.Add($candidate)
            return [pscustomobject]@{ Upn = $candidate; WasFallback = $true; CollisionCount = $collisions.Count; Collisions = [string[]]$collisions }
        }
        $collisions.Add($candidate)
    }
}

# Familiar nouns and simple adjectives, including ordinary plural forms. Keeping
# this vocabulary explicit makes its finite selection space auditable.
$script:StudentPasswordWords = @(
    'Aal', 'Abend', 'Acker', 'Adler', 'Affe', 'Affen', 'Ahorn', 'Allee', 'Alm', 'Ameise', 'Ampel', 'Amsel',
    'Ananas', 'Anker', 'Apfel', 'Arm', 'Arme', 'Ast', 'Aeste', 'Atlas', 'Auge', 'Augen', 'Auto', 'Autos',
    'Baby', 'Bach', 'Backen', 'Bagger', 'Bahn', 'Bahnen', 'Ball', 'Ballett', 'Bambus', 'Banane', 'Bank', 'Bart',
    'Basteln', 'Bauch', 'Bauen', 'Bauer', 'Baum', 'Bauern', 'Becher', 'Becken', 'Beere', 'Beeren', 'Beet', 'Beete',
    'Bein', 'Beine', 'Berg', 'Berge', 'Besen', 'Bett', 'Betten', 'Biber', 'Biene', 'Bienen', 'Bild', 'Bilder',
    'Birke', 'Birne', 'Birnen', 'Blatt', 'Blau', 'Blaue', 'Blech', 'Blick', 'Blitz', 'Blume', 'Blumen', 'Boden',
    'Bogen', 'Bohne', 'Bohnen', 'Boot', 'Boote', 'Bote', 'Box', 'Braten', 'Braun', 'Breit', 'Breite', 'Brett',
    'Brief', 'Brille', 'Brot', 'Brote', 'Brunnen', 'Buch', 'Buchen', 'Bude', 'Bunt', 'Bunte', 'Burg', 'Burgen',
    'Bus', 'Busch', 'Busse', 'Butter', 'Cafe', 'Camping', 'Chor', 'Clown', 'Comic', 'Dach', 'Dachs', 'Dackel',
    'Danke', 'Daten', 'Daumen', 'Decke', 'Decken', 'Delfin', 'Denken', 'Dicht', 'Dicke', 'Dinge', 'Dino', 'Dinos',
    'Distel', 'Dose', 'Dosen', 'Drachen', 'Drehen', 'Dreieck', 'Duft', 'Dunkel', 'Ecke', 'Ecken', 'Efeu', 'Eichel',
    'Eiche', 'Eichen', 'Eier', 'Eimer', 'Eis', 'Elch', 'Elche', 'Elster', 'Eltern', 'Ende',
    'Engel', 'Ente', 'Enten', 'Erbse', 'Erbsen', 'Erde', 'Erden', 'Erika', 'Ernte', 'Ernten', 'Esel', 'Espe',
    'Essen', 'Eule', 'Eulen', 'Fabel', 'Faden', 'Fahne', 'Fahnen', 'Fahrrad', 'Falke', 'Familie', 'Farbe',
    'Farben', 'Farn', 'Feder', 'Federn', 'Fee', 'Feen', 'Feier', 'Feiern', 'Feige', 'Feigen', 'Feld', 'Felder',
    'Fell', 'Ferien', 'Ferkel', 'Ferse', 'Feste', 'Feuer', 'Fichte', 'Figur', 'Film', 'Filme', 'Finger', 'Fink',
    'Fisch', 'Fische', 'Flamme', 'Fleck', 'Fliege', 'Flocke', 'Flora', 'Floh', 'Floss', 'Flug', 'Fluss',
    'Flut', 'Form', 'Formen', 'Foto', 'Fotos', 'Frei', 'Freie', 'Freude', 'Freund', 'Frisch', 'Froh', 'Frohe',
    'Frosch', 'Frucht', 'Fuchs', 'Fuellen', 'Funk', 'Gabel', 'Gans', 'Ganz', 'Garage', 'Garten', 'Gas', 'Gasse',
    'Gassen', 'Geige', 'Gelb', 'Gelbe', 'Gemse', 'Gerste', 'Geste', 'Giraffe', 'Glanz', 'Glas', 'Glatte', 'Gleis',
    'Glocke', 'Globus', 'Gold', 'Graben', 'Gras', 'Grau', 'Graue', 'Greif', 'Grille', 'Grosse', 'Gruen', 'Gruene',
    'Grund', 'Gurke', 'Gurken', 'Haar', 'Haare', 'Hafen', 'Hafer', 'Hahn', 'Hai', 'Hain', 'Haken', 'Halle',
    'Hallo', 'Halm', 'Halme', 'Hals', 'Hamster', 'Hand', 'Hase', 'Hasen', 'Haube', 'Haus', 'Hecke', 'Hecken',
    'Heft', 'Hefte', 'Heide', 'Heim', 'Heimat', 'Heiter', 'Hell', 'Helle', 'Helm', 'Hemd', 'Hemden', 'Herbst',
    'Herd', 'Herde', 'Herz', 'Herzen', 'Heu', 'Hilfe', 'Himmel', 'Hirsch', 'Hobby', 'Hof', 'Holz', 'Honig',
    'Horn', 'Hose', 'Hosen', 'Hub', 'Hund', 'Hunde', 'Hut', 'Husten', 'Igel', 'Imker', 'Insel', 'Inseln',
    'Jacke', 'Jacken', 'Jahr', 'Jahre', 'Jasmin', 'Jubel', 'Judo', 'Juli', 'Junge', 'Juni', 'Kabel', 'Kachel',
    'Kaefer', 'Kaese', 'Kakao', 'Kamel', 'Kamele', 'Kamin', 'Kamm', 'Kanne', 'Kannen', 'Kanu', 'Kapitel', 'Kappe',
    'Karo', 'Karte', 'Karten', 'Karton', 'Kasse', 'Kasten', 'Katze', 'Katzen', 'Kegel', 'Keks', 'Kekse', 'Keller',
    'Kerze', 'Kerzen', 'Kessel', 'Kette', 'Ketten', 'Kiesel', 'Kiefer', 'Kino', 'Kirsche', 'Kiste', 'Kisten', 'Kissen',
    'Kiwi', 'Klang', 'Kleber', 'Klee', 'Kleid', 'Klein', 'Kleine', 'Klecks', 'Klinge', 'Klopfen', 'Klug', 'Kluge',
    'Knall', 'Knete', 'Knopf', 'Koala', 'Koch', 'Kochen', 'Koffer', 'Kohle', 'Kolben', 'Komet', 'Kopf',
    'Korb', 'Korn', 'Kragen', 'Kran', 'Kraut', 'Krebs', 'Kreide', 'Kreis', 'Krone', 'Kronen', 'Krug', 'Krume',
    'Kuchen', 'Kuckuck', 'Kugel', 'Kugeln', 'Kuh', 'Kunst', 'Kurs', 'Kurz', 'Kurze', 'Kuss', 'Kutsche', 'Lachen',
    'Laden', 'Lager', 'Lama', 'Lamm', 'Lampe', 'Lampen', 'Land', 'Lappen', 'Laster', 'Laterne', 'Laub', 'Lauch',
    'Lauf', 'Laufen', 'Laune', 'Laut', 'Laute', 'Leben', 'Leder', 'Lego', 'Lehm', 'Lehrer', 'Leicht', 'Leim',
    'Leine', 'Leise', 'Leiter', 'Lernen', 'Lesen', 'Licht', 'Lied', 'Lieder', 'Lilie', 'Linde', 'Linie', 'Linien',
    'Linse', 'Linsen', 'Liste', 'Listen', 'Lob', 'Loch', 'Locke', 'Locken', 'Loewe', 'Luft', 'Lupe', 'Lupen',
    'Lustig', 'Maler', 'Malen', 'Mama', 'Mandel', 'Mango', 'Mann', 'Mantel', 'Mappe', 'Marke', 'Markt', 'Maske',
    'Masten', 'Maus', 'Meer', 'Mehl', 'Meile', 'Meise', 'Meisen', 'Melone', 'Mensch', 'Messen', 'Milch', 'Minze',
    'Mitte', 'Mode', 'Mohn', 'Moewe', 'Moehre', 'Mond', 'Moos', 'Morgen', 'Motor', 'Motte', 'Motten', 'Muffin',
    'Mulde', 'Mund', 'Murmel', 'Muschel', 'Musik', 'Muster', 'Mut', 'Mutig', 'Mutige', 'Mutter', 'Nadel', 'Nadeln',
    'Nager', 'Namen', 'Nase', 'Nasen', 'Natur', 'Nebel', 'Nelke', 'Nelken', 'Nemo', 'Nest', 'Nester', 'Netz',
    'Netze', 'Neugier', 'Nische', 'Nixe', 'Note', 'Noten', 'Nudel', 'Nudeln', 'Nuss', 'Obst', 'Ochse', 'Ocker',
    'Ofen', 'Ohr', 'Ohren', 'Olive', 'Onkel', 'Orange', 'Orgel', 'Otter', 'Paket', 'Pakete', 'Palast', 'Palme',
    'Palmen', 'Panda', 'Papa', 'Papier', 'Park', 'Parks', 'Party', 'Pasta', 'Pause', 'Pausen', 'Pendel', 'Perle',
    'Perlen', 'Pferd', 'Pferde', 'Pfote', 'Pfoten', 'Piano', 'Pilot', 'Pilz', 'Pilze', 'Pinsel', 'Pirat', 'Pizza',
    'Plakat', 'Plan', 'Platz', 'Pony', 'Post', 'Poster', 'Prise', 'Probe', 'Proben', 'Pudel', 'Pult',
    'Punkt', 'Punkte', 'Puppe', 'Puppen', 'Puzzle', 'Qualle', 'Quark', 'Quelle', 'Rabe', 'Raben', 'Rad', 'Radio',
    'Rakete', 'Rand', 'Rasen', 'Raupe', 'Raupen', 'Raum', 'Rebe', 'Regen', 'Reh', 'Rehe', 'Reifen', 'Reihe',
    'Reim', 'Reime', 'Reis', 'Reise', 'Reiten', 'Rennen', 'Riese', 'Riesen', 'Riff', 'Rinde', 'Ring', 'Ringe',
    'Ritter', 'Robe', 'Roggen', 'Rohr', 'Rohre', 'Rolle', 'Rollen', 'Roman', 'Rose', 'Rosen', 'Rosine', 'Rot',
    'Rote', 'Ruder', 'Rufen', 'Ruhe', 'Ruhig', 'Runde', 'Rutsche', 'Saat', 'Sache', 'Sachen', 'Saft', 'Sage',
    'Saite', 'Salat', 'Salbei', 'Salbe', 'Salz', 'Samen', 'Sand', 'Satin', 'Satz', 'Schaf', 'Schale', 'Scharf',
    'Schatz', 'Schau', 'Schiff', 'Schilf', 'Schnee', 'Schuh', 'Schuhe', 'Schule', 'Schwan', 'Sechs', 'See', 'Segel',
    'Seide', 'Seife', 'Seil', 'Seile', 'Seite', 'Seiten', 'Senf', 'Sessel', 'Sieben', 'Silbe', 'Silber', 'Singen',
    'Sippe', 'Sitz', 'Sitzen', 'Sofa', 'Sommer', 'Sonne', 'Spass', 'Specht', 'Spiel', 'Spiele', 'Spinne', 'Spitze',
    'Sport', 'Sprung', 'Spule', 'Spur', 'Spuren', 'Stall', 'Stamm', 'Stange', 'Staub', 'Steg', 'Stein', 'Steine',
    'Stern', 'Sterne', 'Stift', 'Stille', 'Stock', 'Stoff', 'Stolz', 'Strand', 'Strauch', 'Stroh', 'Strom', 'Stube',
    'Stuhl', 'Stunde', 'Suche', 'Suppe', 'Suppen', 'Tafel', 'Tage', 'Takt', 'Tal', 'Taler', 'Tanne', 'Tannen',
    'Tanz', 'Tanzen', 'Tasse', 'Tassen', 'Tau', 'Taube', 'Tauben', 'Teich', 'Teil', 'Teile', 'Teller', 'Tennis',
    'Text', 'Tier', 'Tiere', 'Tiger', 'Tinte', 'Tisch', 'Titel', 'Toast', 'Tofu', 'Tomate', 'Ton', 'Topf',
    'Tor', 'Torte', 'Torten', 'Traube', 'Traum', 'Treff', 'Treppe', 'Treu', 'Treue', 'Trommel', 'Tuch', 'Tulpe',
    'Tulpen', 'Tunnel', 'Turm', 'Turnen', 'Ufer', 'Uhr', 'Uhren', 'Uhu', 'Ulme', 'Urlaub', 'Vase', 'Vasen',
    'Vater', 'Verein', 'Vogel', 'Wabe', 'Waben', 'Wache', 'Wachs', 'Wagen', 'Wahl', 'Wal', 'Wald', 'Wand',
    'Wange', 'Wangen', 'Wanne', 'Wannen', 'Warm', 'Warme', 'Wasser', 'Watten', 'Weg', 'Wege', 'Weide', 'Weiden',
    'Weiler', 'Weise', 'Weiss', 'Weisse', 'Weit', 'Weite', 'Weizen', 'Welle', 'Wellen', 'Welt', 'Wende', 'Werfen',
    'Werk', 'Wert', 'Wespe', 'Wetter', 'Wiese', 'Wiesen', 'Wild', 'Wilde', 'Wille', 'Wind', 'Winde', 'Winkel',
    'Winter', 'Wippe', 'Wissen', 'Witz', 'Witze', 'Woche', 'Wolf', 'Wolke', 'Wolken', 'Wolle', 'Wort', 'Worte',
    'Wunder', 'Wunsch', 'Wurm', 'Wurzel', 'Zahl', 'Zahlen', 'Zahn', 'Zauber', 'Zaun', 'Zebra', 'Zehen', 'Zeiger',
    'Zeit', 'Zelt', 'Zelte', 'Zettel', 'Zeug', 'Ziege', 'Ziel', 'Zimmer', 'Zimt', 'Zopf', 'Zug', 'Zweck', 'Zweig'
)

function Get-StudentPasswordPrefixes {
    if ($null -eq (Get-Variable -Name StudentPasswordPrefixes -Scope Script -ErrorAction SilentlyContinue)) {
        $prefixes = [Collections.Generic.List[string]]::new()
        foreach ($left in $script:StudentPasswordWords) {
            foreach ($right in $script:StudentPasswordWords) {
                if (($left.Length + $right.Length) -eq 10) { $prefixes.Add("$left$right") }
            }
        }
        $script:StudentPasswordPrefixes = $prefixes.ToArray()
    }
    # Return the cached array as one object, avoiding a 200,000-item pipeline per password.
    return ,$script:StudentPasswordPrefixes
}

function New-StudentPassword {
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][System.Collections.Generic.HashSet[string]] $UsedPasswords,
        [scriptblock] $RandomIndexScriptBlock
    )
    $prefixes = Get-StudentPasswordPrefixes
    for ($attempt = 0; $attempt -lt 100; $attempt++) {
        if ($null -eq $RandomIndexScriptBlock) {
            $prefixIndex = [Security.Cryptography.RandomNumberGenerator]::GetInt32(0, $prefixes.Count)
        } else {
            $prefixIndex = & $RandomIndexScriptBlock $prefixes.Count
            if ($prefixIndex -isnot [int] -or $prefixIndex -lt 0 -or $prefixIndex -ge $prefixes.Count) {
                throw 'Der injizierte Zufallsindex muss eine ganze Zahl innerhalb der Wortauswahl sein.'
            }
        }
        $number = [Security.Cryptography.RandomNumberGenerator]::GetInt32(10, 100)
        $password = '{0}{1:00}' -f $prefixes[$prefixIndex], $number
        if ($UsedPasswords.Add($password)) { return $password }
    }
    throw 'Nach 100 Versuchen konnte kein eindeutiges Schülerpasswort erzeugt werden.'
}
