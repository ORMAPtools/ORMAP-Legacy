# ---------------------------------------------------------------------------
# OrmapLayersConfig.py
# Created by: Shad Campbell
# Date: 3/11/2011
# Updated by: 
# Description: This is a configuration file to be customized by each county.
# Do not delete any of the items in this file.  If they are not in use then
# specify thier value and/or definition query  to "".
# ---------------------------------------------------------------------------


LOTSANNO_LAYER="LotsAnno"
LOTSANNO_QD="\"MapNumber\" = '*MapNumber*'OR \"MapNumber\" is NULL OR \"MapNumber\" = ''"

PLATSANNO_LAYER="PlatsAnno"
PLATSANNO_QD="\"MapNumber\" = '*MapNumber*' OR \"MapNumber\" is NULL OR \"MapNumber\" = ''"

TAXCODEANNO_LAYER="TaxCodeAnno"
TAXCODEANNO_QD="\"MapNumber\" = '*MapNumber*' OR \"MapNumber\" is NULL OR \"MapNumber\" = ''"

TAXNUMANNO_LAYER="TaxlotNumberAnno"
TAXNUMANNO_QD="\"MapNumber\" = '*MapNumber*' OR \"MapNumber\" is NULL OR \"MapNumber\" = ''"

ACRESANNO_LAYER="TaxlotAcresAnno"
ACRESANNO_QD="\"MapNumber\" = '*MapNumber*' OR \"MapNumber\" is NULL OR \"MapNumber\" = ''"

ANNO10_LAYER="Anno0010scale"
ANNO10_QD="\"MapNumber\" = '*MapNumber*' OR \"MapNumber\" is NULL OR \"MapNumber\" = ''"

ANNO20_LAYER="Anno0020scale"
ANNO20_QD="\"MapNumber\" = '*MapNumber*' OR \"MapNumber\" is NULL OR \"MapNumber\" = ''"

ANNO30_LAYER="Anno0030scale"
ANNO30_QD="\"MapNumber\" = '*MapNumber*' OR \"MapNumber\" is NULL OR \"MapNumber\" = ''"

ANNO40_LAYER="Anno0040scale"
ANNO40_QD="\"MapNumber\" = '*MapNumber*' OR \"MapNumber\" is NULL OR \"MapNumber\" = ''"

ANNO50_LAYER="Anno0050scale"
ANNO50_QD="\"MapNumber\" = '*MapNumber*' OR \"MapNumber\" is NULL OR \"MapNumber\" = ''"

ANNO60_LAYER="Anno0060scale"
ANNO60_QD="\"MapNumber\" = '*MapNumber*' OR \"MapNumber\" is NULL OR \"MapNumber\" = ''"

ANNO100_LAYER="Anno0100scale"
ANNO100_QD="\"MapNumber\" = '*MapNumber*' OR \"MapNumber\" is NULL OR \"MapNumber\" = ''"

ANNO200_LAYER="Anno0200scale"
ANNO200_QD="\"MapNumber\" = '*MapNumber*' OR \"MapNumber\" is NULL OR \"MapNumber\" = ''"

ANNO400_LAYER="Anno0400scale"
ANNO400_QD="\"MapNumber\" = '*MapNumber*' OR \"MapNumber\" is NULL OR \"MapNumber\" = ''"

ANNO800_LAYER="Anno0800scale"
ANNO800_QD="\"MapNumber\" = '*MapNumber*' OR \"MapNumber\" is NULL OR \"MapNumber\" = ''"

ANNO2000_LAYER="Anno2000scale"
ANNO2000_QD="\"MapNumber\"='*MapNumber*' OR \"MapNumber\" is NULL OR \"MapNumber\" = ''"

CORNER_ABOVE_LAYER="Corner"
CORNER_ABOVE_QD="\"MapNumber\"='*MapNumber*'"

TAXCODELINES_ABOVE_LAYER="TaxCodeLines - Above"
TAXCODELINES_ABOVE_QD=""

TAXLOTLINES_ABOVE_LAYER="TaxlotLines - Above"
TAXLOTLINES_ABOVE_QD="\"LineType\" <> 32"

REFLINES_ABOVE_LAYER="ReferenceLines - Above"
REFLINES_ABOVE_QD="\"MAPNUMBER\" = '*MapNumber*'"

CARTOLINES_ABOVE_LAYER="CartographicLines - Above"
CARTOLINES_ABOVE_QD=""

WATERLINES_ABOVE_LAYER="WaterLines - Above"
WATERLINES_ABOVE_QD=""

WATER_ABOVE_LAYER="Water - Above"
WATER_ABOVE_QD=""

MAPINDEXSEEMAP_LAYER=""
MAPINDEXSEEMAP_QD=""

MAPINDEX_LAYER="SeeMaps"
MAPINDEX_QD="\"IndexMap\" = '*MapNumber*'"

CORNER_BELOW_LAYER="Corner - Below"
CORNER_BELOW_QD=""

TAXCODELINES_BELOW_LAYER="TaxCodeLines - Below"
TAXCODELINES_BELOW_QD=""

TAXLOTLINES_BELOW_LAYER="TaxlotLines - Below"
TAXLOTLINES_BELOW_QD=""

REFLINES_BELOW_LAYER="ReferenceLines - Below"
REFLINES_BELOW_QD=""

CARTOLINES_BELOW_LAYER="CartographicLines - Below"
CARTOLINES_BELOW_QD=""

WATERLINES_BELOW_LAYER="WaterLines - Below"
WATERLINES_BELOW_QD=""

WATER_BELOW_LAYER="Water - Below"
WATER_BELOW_QD=""

PAGELAYOUT_TABLE="giscarto.CREATOR_ASR.PAGELAYOUTELEMENTS"
CANCELLEDNUMBERS_TABLE="giscarto.CREATOR_ASR.CANCELLEDNUMBERS"
CUSTOMDEFINITIONQUERIES_TABLE="CustomDefinitionQueries"

EXTRA1_LAYER="Arrow0010scale"
EXTRA1_QD="\"MapNumber\"='*MapNumber*' OR \"MapNumber\" is NULL OR \"MapNumber\" = ''"

EXTRA2_LAYER="Arrow0020scale"
EXTRA2_QD="\"MapNumber\"='*MapNumber*' OR \"MapNumber\" is NULL OR \"MapNumber\" = ''"

EXTRA3_LAYER="Arrow0030scale"
EXTRA3_QD="\"MapNumber\"='*MapNumber*' OR \"MapNumber\" is NULL OR \"MapNumber\" = ''"

EXTRA4_LAYER="Arrow0040scale"
EXTRA4_QD="\"MapNumber\"='*MapNumber*' OR \"MapNumber\" is NULL OR \"MapNumber\" = ''"

EXTRA5_LAYER="Arrow0050scale"
EXTRA5_QD="\"MapNumber\"='*MapNumber*' OR \"MapNumber\" is NULL OR \"MapNumber\" = ''"

EXTRA6_LAYER="Arrow0100scale"
EXTRA6_QD="\"MapNumber\"='*MapNumber*' OR \"MapNumber\" is NULL OR \"MapNumber\" = ''"

EXTRA7_LAYER="Arrow0200scale"
EXTRA7_QD="\"MapNumber\"='*MapNumber*' OR \"MapNumber\" is NULL OR \"MapNumber\" = ''"

EXTRA8_LAYER="Arrow0400scale"
EXTRA8_QD="\"MapNumber\"='*MapNumber*' OR \"MapNumber\" is NULL OR \"MapNumber\" = ''"

EXTRA9_LAYER="Arrow2000scale"
EXTRA9_QD="\"MapNumber\"='*MapNumber*' OR \"MapNumber\" is NULL OR \"MapNumber\" = ''"

EXTRA10_LAYER="MapSecLines - Below"
EXTRA10_QD="\"MapNumber\"='*MapNumber*'"

EXTRA11_LAYER="Railroad"
EXTRA11_QD="CL <> 'Y'"

EXTRA12_LAYER="MapArea"
EXTRA12_QD="\"MapNumber\"='*MapNumber*'"

EXTRA13_LAYER=""
EXTRA13_QD=""

EXTRA14_LAYER="Taxlots - Above"
EXTRA14_QD="\"MapNumber\"='*MapNumber*'"

EXTRA15_LAYER="Arrow0060scale"
EXTRA15_QD="\"MapNumber\"='*MapNumber*' OR \"MapNumber\" is NULL OR \"MapNumber\" = ''"

EXTRA16_LAYER="Landmarks"
EXTRA16_QD="\"MapNumber\"='*MapNumber*' OR \"MapNumber\" is NULL OR \"MapNumber\" = ''"

EXTRA17_LAYER=""
EXTRA17_QD=""

EXTRA18_LAYER=""
EXTRA18_QD=""

EXTRA19_LAYER=""
EXTRA19_QD=""

EXTRA20_LAYER=""
EXTRA20_QD=""
