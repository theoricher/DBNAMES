###############################################################################
## OCSINVENTORY-NG
## Copyleft Théo RICHER 2019
## Web : https://www.ocsinventory-ng.org
##
## This code is open source and may be copied and modified as long as the source
## code is always made freely available.
## Please refer to the General Public Licence http://www.gnu.org/ or Licence.txt
################################################################################
 
package Apache::Ocsinventory::Plugins::Dbnames::Map;
 
use strict;
 
use Apache::Ocsinventory::Map;
#Plugin DBNAMES
$DATA_MAP{dbnames} = {
mask => 0,
multi => 1,
auto => 1,
delOnReplace => 1,
sortBy => 'NAME',
writeDiff => 1,
cache => 0,
mandatory => 1,
fields => {
     NAME => {},
	 FILE => {},
}
 
};
1;
