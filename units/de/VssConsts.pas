(* Delphi-Unit
   Resource strings for VssUtils - German
   ======================================
   
   © Dr. J. Rathlev, D-24222 Schwentinental (kontakt(a)rathlev-home.de)

   The contents of this file may be used under the terms of the
   Mozilla Public License ("MPL") or
   GNU Lesser General Public License Version 2 or later (the "LGPL")

   Software distributed under this License is distributed on an "AS IS" basis,
   WITHOUT WARRANTY OF ANY KIND, either express or implied. See the License for
   the specific language governing rights and limitations under the License.

   Version 1.0: February 2015
           2.1: January 2016 - changed to "resourcestring"
   last updated: Feb. 2017
   *)

unit VssConsts;

interface

resourcestring
  rsVssNotAvail = 'Die Volumen-Schattenkopie konnte auf diesem System nicht initialisiert werden!';
  rsInitBackupComps = 'Das IVssBackupComponents-Interface wird initialisiert ...';
  rsInitMetaData = 'Writer-Metadaten werden initialisiert ...';
  rsGatherMetaData = 'Writer-Metadaten werden erfasst ...';
  rsGatherWriterStatus = 'Status der Writer wird erfasst ...';
  rsDscvDirExclComps = 'Direkt auszuschließende Komponenten werden ermittelt ...';
  rsDscvNsExclComps = 'Komponenten, die nicht zum Schattenkopiesatz gehören, werden ermittelt ...';
  rsDscvAllExclComps = 'Die ausgeschlossenen Komponenten werden ermittelt ...';
  rsDscvExclWriters = 'Ausgeschlossene Writer werden ermittelt ...';
  rsDscvExpInclComps ='Ausdrücklich einzuschließende Komponenten werden ermittelt ...';
  rsSelExpInclComps = 'Ausdrücklich einzuschließende Komponenten auswählen ...';
  rsSaveBackCompsDoc = 'Das Backupkomponenten-Dokument wird gespeichert ...';
  rsAddVolSnapSet = 'Dem Snapshotsatz werden Datenträger hinzugefügt ...';
  rsPrepBackup = 'Das Backup wird vorbereitet ...';
  rsCreateShadowCopy = 'Die Schattenkopie wird erstellt ...';
  rsCreateCopySuccess = 'Der Schattenkopiesatz wurde erfolgreich erstellt';
  rsQueryShadowCopies = 'Alle Schattenkopien im System werden abgefragt ...';
  rsQuerySnapshotSetID = 'Alle Schattenkopien mit dem SnapshotSetID %s werden abgefragt ...';
  rsCreateShadowSet = 'Ein Schattenkopiesatz {%s} wird erstellt ...';
  rsCompleteBackup = 'Das Backup wird abgeschlossen ...';

implementation

end.
