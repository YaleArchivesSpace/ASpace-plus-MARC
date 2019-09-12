# ArchivesSpace MARC to Voyager

For this process, we have decided to take a two-pronged approach.  

First, we are using a modified version of NYU's MARCXML export plugin in ArchivesSpace.  See: https://github.com/YaleArchivesSpace/yale_marcxml_export_plugin With this plugin, we ensure that all of the data that we need to be serialized to the MARCXML file does get serialized.  For instance, by default, ArchivesSpace's MARCXML exporter does *not* include the "EAD ID" field.  However, we rely on this field to be in our MARC records, so we use this plugin to make sure that happens. Ditto for additional data that would not be included in the exports by default, such as information about the top container records and their restrictions.

Second, we use XSLT to transform the MARCXML files that are produced by ArchiveSpace. At this stage, we ensure that all of the other local transformations happen -- as long as the data is in the first export, we can transform it and update it as needed during this step. We have decided to keep these additional transformations outside of the ArchivesSpace plugin so that we can more easily and quickly adjust those changes as needed.
