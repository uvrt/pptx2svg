"""TEMPORARY (s3): this repository's catalogue of built-in table styles is ooxml-common's.

Run once on the pull request that replaces the copy in ``parse/table_styles_builtin.py``
with ``ooxml_common.drawingml.table_styles``; removed with the copy.
"""

from ooxml_common.drawingml import table_styles as shared

from pptx2svg.parse import table_styles_builtin as copy


def test_the_copy_is_the_shared_catalogue():
    assert copy.BUILTIN_TABLE_STYLES == shared.BUILTIN_TABLE_STYLES
    assert list(copy.BUILTIN_TABLE_STYLES) == list(shared.BUILTIN_TABLE_STYLES)
    assert copy._MEASURED_BORDERS == shared.MEASURED_BORDERS
    assert copy._THEMED_STYLE_XML == shared.THEMED_STYLE_XML
    assert copy._THEMED_STYLES == shared.THEMED_STYLES
    assert [name for name, _ in copy.TABLE_STYLE_REGIONS] == list(shared.REGIONS)
