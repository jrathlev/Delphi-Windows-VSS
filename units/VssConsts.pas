(* Delphi-Unit
   Resource strings for VssUtils - English
   =======================================

   © Dr. J. Rathlev, D-24222 Schwentinental (kontakt(a)rathlev-home.de)

   The contents of this file may be used under the terms of the
   Mozilla Public License ("MPL") or
   GNU Lesser General Public License Version 2 or later (the "LGPL")

   Software distributed under this License is distributed on an "AS IS" basis,
   WITHOUT WARRANTY OF ANY KIND, either express or implied. See the License for
   the specific language governing rights and limitations under the License.

   Version 1.0: February 2015
           2.1: January 2016 - changes to use "resourcestring" for localization
   *)

unit VssConsts;

interface

resourcestring
  rsVssNotAvail = 'Volume Shadow Copy could not be initialized on this system!';
  rsInitBackupComps = 'Initializing IVssBackupComponents Interface ...';
  rsInitMetaData = 'Initialize writer metadata ...';
  rsGatherMetaData = 'Gathering writer metadata ...';
  rsGatherWriterStatus = 'Gathering writer status ...';
  rsDscvDirExclComps = 'Discover directly excluded components ...';
  rsDscvNsExclComps = 'Discover components that reside outside the shadow set ...';
  rsDscvAllExclComps = 'Discover all excluded components ...';
  rsDscvExclWriters = 'Discover excluded writers ...';
  rsDscvExpInclComps = 'Discover explicitly included components ...';
  rsSelExpInclComps = 'Select explicitly included components ...';
  rsSaveBackCompsDoc = 'Saving the backup components document ...';
  rsAddVolSnapSet = 'Add volumes to snapshot set ...';
  rsPrepBackup = 'Preparing for backup ...';
  rsCreateShadowCopy = 'Creating the shadow copy ...';
  rsCreateCopySuccess = 'Shadow copy set successfully created';
  rsQueryShadowCopies = 'Querying all shadow copies in the system ...';
  rsQuerySnapshotSetID = 'Querying all shadow copies with the SnapshotSetID %s ...';
  rsCreateShadowSet = 'Creating shadow set {%s} ...';
  rsCompleteBackup = 'Completing the backup ...';

implementation

end.
