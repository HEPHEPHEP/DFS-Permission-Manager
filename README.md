# DFS-Permission-Manager

## Beschreibung

Dieses PowerShell-Skript ist ein grafisches Tool zur Verwaltung von Ordnern, DFS-Verknüpfungen (Distributed File System) und NTFS-Berechtigungen in einer Windows-Active-Directory-Umgebung. Es ermöglicht die Erstellung neuer Ordnerstrukturen, die Generierung entsprechender Active-Directory-Gruppen für Berechtigungen, das Setzen von NTFS-Rechten und die Verwaltung von Benutzern in Gruppen.

Das Tool ist in Tabs unterteilt:
- **Ordner erstellen**: Erstellen von Ordnern, Unterordnern, DFS-Links und AD-Gruppen.
- **Berechtigungen**: Anzeige und Zuweisung von Berechtigungen auf bestehenden Ordnern.
- **Benutzer zu Gruppen**: Suche und Hinzufügen von Benutzern zu Berechtigungsgruppen.

## Anforderungen

- PowerShell Version 5.1 oder höher
- Active Directory PowerShell-Modul (Installiere mit `Install-Module -Name ActiveDirectory`, falls nicht vorhanden)
- Administratorrechte auf dem Dateiserver und im Active Directory
- Zugriff auf DFS-Namespace und Dateiserver
- .NET Assemblies: PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms (standardmäßig in Windows verfügbar)

## Installation

1. Klone das Repository oder lade das Skript herunter:

   ```git clone https://github.com/HEPHEPHEP/DFS-Permission-Manager.git```

2. Navigiere in das Verzeichnis:

   ```cd DFS-Permission-Manager```

## Verwendung

Führe das Skript in einer PowerShell-Konsole mit Administratorrechten aus:

```.\DFS-Permission-Manager.ps1```

Die grafische Benutzeroberfläche (GUI) basierend auf WPF öffnet sich. Konfiguriere die Einstellungen in der Config-Sektion des Skripts, falls nötig (z.B. DFS-Root, Gruppen-OU).

### Konfiguration

Im Skript gibt es eine `$Config`-HashTable, die angepasst werden kann:
- `GroupNameSchema`: Schema für Gruppennamen (z.B. "FS_{FolderName}_{Permission}").
- `GroupDescriptionSchema`: Beschreibung für Gruppen.
- `GroupOU`: OU im Active Directory, wo Gruppen erstellt werden.
- `DFSRoot`: DFS-Root-Pfad.
- `DefaultFileServer`: Standard-Dateiserver.
- `PermissionTypes`: Definierte Berechtigungstypen (Lesen, Ändern, Vollzugriff usw.) mit zugehörigen NTFS-Rechten.
- `PermissionLabels`: Kürzel für Berechtigungen (z.B. "RO" für Read-Only).

Beispiele für Berechtigungstypen:
- **Lesen**: ReadAndExecute, vererbt.
- **Ändern**: Modify, vererbt.
- **Vollzugriff**: FullControl, vererbt.
- **Auflisten**: ReadAndExecute, nicht vererbt.
- **NurDieserOrdner**: Modify, nicht vererbt.

### Funktionen

- **Ordner erstellen**:
  - Gib Basisverzeichnis, neuen Ordnernamen und optionale Unterordner an.
  - Optionen: DFS erstellen, Gruppen erstellen, Vererbung deaktivieren.
  - Vorschau der zu erstellenden Elemente.

- **Berechtigungen**:
  - Lade Ordnerstruktur in einen TreeView.
  - Zeige aktuelle NTFS-Berechtigungen in einer DataGrid.
  - Erstelle neue Gruppen und weise Berechtigungen zu.
  - Option: Traverse-Rechte für Elternordner setzen.

- **Benutzer zu Gruppen**:
  - Suche nach Benutzern im Active Directory.
  - Füge Benutzer zu ausgewählten Berechtigungsgruppen hinzu.

## Beispiele

1. **Ordner erstellen**:
   - Basisverzeichnis: `\\fileserver01\share`
   - Neuer Ordner: `ProjektX`
   - Unterordner: `Dokumente\nArchiv`
   - Aktiviere DFS und Gruppen: Erstellt DFS-Link, AD-Gruppen wie `FS_ProjektX_RO` und setzt NTFS-Rechte.

2. **Berechtigungen setzen**:
   - Wähle Ordner aus dem TreeView.
   - Wähle Berechtigungstyp (z.B. "Ändern").
   - Klicke "Gruppe erstellen": Erstellt Gruppe und weist Rechte zu.

## Hinweise

- Das Skript verwendet WPF für die GUI, was eine moderne Windows-Umgebung voraussetzt.
- Stelle sicher, dass der Dateiserver und DFS-Namespace korrekt konfiguriert sind.
- Teste in einer Testumgebung, da Änderungen an Berechtigungen und AD irreversibel sein können.
- Fehlermeldungen werden in der Konsole ausgegeben.

## Lizenz

MIT License



