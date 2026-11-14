"""Pure ribbon XML edits: move a control, group controls into a menu, image links."""

from __future__ import annotations

import pytest
from conftest import RIBBON

from vba_addin_editor.adapters.xml_codec import validate_xml_bytes
from vba_addin_editor.services import ribbon_service as rs


def _at(text: str, needle: str) -> int:
    return text.index(needle) + 2


def _group(text: str, group_id: str) -> str:
    start = text.index(f'<group id="{group_id}"')
    return text[start : text.index("</group>", start) + len("</group>")]


def test_move_swaps_siblings_and_keeps_comments_in_place():
    text, moved = rs.move_element(RIBBON, _at(RIBBON, 'id="btnMissing"'), -1)
    assert _group(text, "grpMain") == (
        '<group id="grpMain" label="Main">\n'
        '          <button id="btnMissing" label="Broken" onAction="NoSuchMacro" />\n'
        '          <!-- <button id="btnOld" onAction="Ghost" /> -->\n'
        '          <button id="btnHello" label="Say Hello" size="large" onAction="Hello" />\n'
        "        </group>"
    )
    assert moved.attr("id") == "btnMissing"
    assert text[moved.start : moved.end].startswith('<button id="btnMissing"')
    # Moving back down restores the original text exactly.
    back, _ = rs.move_element(text, moved.start, 1)
    assert back == RIBBON
    validate_xml_bytes(text.encode("utf-8"), part_path="customUI/customUI14.xml")


def test_move_whole_groups_and_stops_at_the_ends():
    text, moved = rs.move_element(RIBBON, _at(RIBBON, 'id="grpEmpty"'), -1)
    assert text.index('id="grpEmpty"') < text.index('id="grpMenus"')
    assert moved.name == "group"
    with pytest.raises(rs.RibbonEditError, match="already the first"):
        rs.move_element(RIBBON, _at(RIBBON, 'id="btnHello"'), -1)
    with pytest.raises(rs.RibbonEditError, match="already the last"):
        rs.move_element(RIBBON, _at(RIBBON, 'id="tglX"'), 1)


def test_move_refuses_fixed_order_elements():
    xml = (
        '<customUI xmlns="x"><ribbon><tabs><tab id="t"><group id="g">\n'
        '  <button id="a" label="A" />\n'
        '  <dialogBoxLauncher><button id="dl" /></dialogBoxLauncher>\n'
        "</group></tab></tabs></ribbon></customUI>"
    )
    with pytest.raises(rs.RibbonEditError, match="dialog box launcher"):
        rs.move_element(xml, _at(xml, 'id="a"'), 1)
    with pytest.raises(rs.RibbonEditError, match="dialog box launcher"):
        rs.move_element(xml, _at(xml, "dialogBoxLauncher"), -1)
    with pytest.raises(rs.RibbonEditError, match="fixed position"):
        rs.move_element(xml, _at(xml, "<ribbon"), 1)
    with pytest.raises(rs.RibbonEditError, match="cursor"):
        rs.move_element(xml, len(xml), 1)


def test_cursor_inside_split_button_moves_the_split_button():
    xml = (
        '<group id="g">\n'
        '  <button id="a" />\n'
        '  <splitButton id="s"><button id="sb" /><menu id="sm"><button id="x" /></menu></splitButton>\n'
        "</group>"
    )
    text, moved = rs.move_element(xml, _at(xml, 'id="sb"'), -1)
    assert moved.name == "splitButton"
    assert text.index('id="s"') < text.index('id="a"')


def test_group_selected_buttons_into_menu():
    start = RIBBON.index('<button id="btnHello"')
    end = RIBBON.index("NoSuchMacro")
    text, offset, length = rs.group_into_menu(
        RIBBON, start, end, control_id="mnuTools", label="Tools & More", size="large"
    )
    assert _group(text, "grpMain") == (
        '<group id="grpMain" label="Main">\n'
        '          <menu id="mnuTools" label="Tools &amp; More" size="large">\n'
        '            <button id="btnHello" label="Say Hello" onAction="Hello" />\n'
        '            <!-- <button id="btnOld" onAction="Ghost" /> -->\n'
        '            <button id="btnMissing" label="Broken" onAction="NoSuchMacro" />\n'
        "          </menu>\n"
        "        </group>"
    )
    assert text[offset : offset + length] == '<menu id="mnuTools" label="Tools &amp; More" size="large">'
    assert not validate_xml_bytes(text.encode("utf-8"), part_path="customUI/customUI14.xml")
    # The new menu is a container the Add Button/Add Menu forms can target.
    assert "mnuTools" in {c.control_id for c in rs.find_containers(text, "p")}


def test_group_inside_a_menu_drops_size_and_cursor_selects_one_control():
    offset = _at(RIBBON, 'id="tglX"')
    text, _o, _l = rs.group_into_menu(
        RIBBON, offset, offset, control_id="mnuSub", label="Sub", size="large", item_size="large"
    )
    # A menu inside a menu takes no size of its own.
    assert '<menu id="mnuSub" label="Sub" itemSize="large">' in text
    assert text.index('id="mnuSub"') < text.index('id="tglX"') < text.index("</menu>", text.index('id="mnuSub"'))


def test_group_refuses_controls_a_menu_cannot_hold():
    xml = '<group id="g">\n  <button id="a" />\n  <editBox id="e" />\n</group>'
    with pytest.raises(rs.RibbonEditError, match="editBox"):
        rs.group_into_menu(xml, xml.index("<button"), xml.index("/>", xml.index("<editBox")), control_id="m", label="M")
    with pytest.raises(rs.RibbonEditError, match="only go in"):
        rs.group_into_menu(RIBBON, RIBBON.index('<group id="grpMain"'), RIBBON.index('<group id="grpEmpty"'), control_id="m", label="M")


def test_relationships_and_content_types():
    assert rs.relationships_path("customUI/customUI14.xml") == "customUI/_rels/customUI14.xml.rels"
    text, rid = rs.add_relationship(rs.new_relationships_text(), rs.IMAGE_REL_TYPE, "images/a.png")
    text, rid2 = rs.add_relationship(text, rs.IMAGE_REL_TYPE, "images/b.png")
    assert (rid, rid2) == ("rIdImg1", "rIdImg2")
    assert rs.relationships(text) == [
        ("rIdImg1", rs.IMAGE_REL_TYPE, "images/a.png"),
        ("rIdImg2", rs.IMAGE_REL_TYPE, "images/b.png"),
    ]
    assert not validate_xml_bytes(text.encode("utf-8"), part_path="customUI/_rels/customUI14.xml.rels")
    one_line = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>'
    filled, _rid = rs.add_relationship(one_line, rs.IMAGE_REL_TYPE, "images/c.png")
    assert filled.endswith('Target="images/c.png"/></Relationships>')

    types = '<Types xmlns="t"><Default Extension="xml" ContentType="application/xml"/></Types>'
    added = rs.ensure_default_content_type(types, "png", "image/png")
    assert '<Default Extension="png" ContentType="image/png"/>' in added
    assert rs.ensure_default_content_type(added, "PNG", "image/png") == added


def test_set_control_image_replaces_imagemso():
    xml = '<group id="g"><button id="b" imageMso="Copy" onAction="X"/></group>'
    element = rs.image_control_at(xml, _at(xml, 'id="b"'))
    assert element is not None and element.name == "button"
    assert rs.set_control_image(xml, element, "rIdImg1") == (
        '<group id="g"><button id="b" onAction="X" image="rIdImg1"/></group>'
    )


# -- designer edits ----------------------------------------------------------------------


def _element(text: str, control_id: str) -> rs.RibbonElement:
    return next(e for e in rs.element_tree(text) if e.attr("id") == control_id)


def _valid(text: str) -> None:
    assert not validate_xml_bytes(text.encode("utf-8"), part_path="customUI/customUI14.xml")


def test_set_element_attrs_sets_adds_and_removes():
    text = rs.set_element_attrs(
        RIBBON, _element(RIBBON, "btnHello"), {"label": "Hi & Bye", "size": None, "screentip": "Greets"}
    )
    assert '<button id="btnHello" label="Hi &amp; Bye" onAction="Hello" screentip="Greets" />' in text
    assert text.count("\n") == RIBBON.count("\n")
    _valid(text)


def test_delete_element_takes_its_lines_and_keeps_the_rest():
    text = rs.delete_element(RIBBON, _element(RIBBON, "btnMissing"))
    assert _group(text, "grpMain") == (
        '<group id="grpMain" label="Main">\n'
        '          <button id="btnHello" label="Say Hello" size="large" onAction="Hello" />\n'
        '          <!-- <button id="btnOld" onAction="Ghost" /> -->\n'
        "        </group>"
    )
    whole = rs.delete_element(RIBBON, _element(RIBBON, "grpMenus"))
    assert "mnuMore" not in whole and 'id="grpEmpty"' in whole
    inline = '<group id="g"><button id="a"/><button id="b"/></group>'
    assert rs.delete_element(inline, _element(inline, "a")) == '<group id="g"><button id="b"/></group>'
    _valid(text)
    _valid(whole)


def test_insert_child_appends_or_goes_beside_a_sibling():
    button = '<button id="btnNew" label="New" />'
    text, offset, length = rs.insert_child(RIBBON, _element(RIBBON, "grpMain"), button)
    assert text[offset : offset + length] == button
    assert _group(text, "grpMain").endswith('onAction="NoSuchMacro" />\n          ' + button + "\n        </group>")
    after, offset, _n = rs.insert_child(
        RIBBON, _element(RIBBON, "grpMain"), button, after=_element(RIBBON, "btnHello")
    )
    assert 'onAction="Hello" />\n          ' + button + "\n          <!--" in after
    assert after[offset:].startswith(button)
    before, offset, _n = rs.insert_child(
        RIBBON, _element(RIBBON, "grpMain"), button, before=_element(RIBBON, "btnHello")
    )
    assert '<group id="grpMain" label="Main">\n          ' + button + '\n          <button id="btnHello"' in before
    assert before[offset:].startswith(button)
    for result in (text, after, before):
        _valid(result)


def test_insert_child_into_empty_and_self_closing_containers():
    menu = rs.build_menu_xml(control_id="mnuNew", label="New", indent="", unit="")
    text, offset, length = rs.insert_child(RIBBON, _element(RIBBON, "grpEmpty"), menu)
    assert (
        '<group id="grpEmpty" label="Empty">\n'
        '          <menu id="mnuNew" label="New">\n'
        "          </menu>\n"
        "        </group>"
    ) in text
    assert text[offset : offset + length].startswith('<menu id="mnuNew"')
    closed = '<tab id="t">\n  <group id="g" label="G"/>\n</tab>'
    opened, offset, _n = rs.insert_child(closed, _element(closed, "g"), '<button id="b" />')
    assert opened == '<tab id="t">\n  <group id="g" label="G">\n    <button id="b" />\n  </group>\n</tab>'
    assert opened[offset:].startswith('<button id="b" />')
    group = rs.build_group_xml(control_id="grpNew", label="New")
    tabbed, _o, _n = rs.insert_child(RIBBON, _element(RIBBON, "tabTools"), group)
    assert '        <group id="grpNew" label="New">\n        </group>\n      </tab>' in tabbed
    _valid(text)
    _valid(tabbed)


def test_insert_child_stays_ahead_of_a_dialog_box_launcher():
    xml = (
        '<group id="g" label="G">\n'
        '  <button id="a" />\n'
        '  <dialogBoxLauncher><button id="dlg" /></dialogBoxLauncher>\n'
        "</group>"
    )
    text, _o, _n = rs.insert_child(xml, _element(xml, "g"), '<button id="b" />')
    assert text.index('id="b"') < text.index("<dialogBoxLauncher>")


def test_relocate_moves_into_a_menu_and_drops_size():
    text, moved = rs.relocate_element(RIBBON, _element(RIBBON, "btnHello"), _element(RIBBON, "mnuMore"))
    assert _group(text, "grpMenus") == (
        '<group id="grpMenus" label="Menus">\n'
        '          <menu id="mnuMore" label="More">\n'
        '            <toggleButton id="tglX" label="Toggle" onAction="OnToggle" getPressed="GetToggle" />\n'
        '            <button id="btnHello" label="Say Hello" onAction="Hello" />\n'
        "          </menu>\n"
        "        </group>"
    )
    assert moved.attr("id") == "btnHello" and 'id="btnHello"' not in _group(text, "grpMain")
    _valid(text)


def test_relocate_moves_a_whole_menu_and_reindents_it():
    text, moved = rs.relocate_element(
        RIBBON, _element(RIBBON, "mnuMore"), _element(RIBBON, "grpMain"), before=_element(RIBBON, "btnHello")
    )
    assert _group(text, "grpMain").startswith(
        '<group id="grpMain" label="Main">\n'
        '          <menu id="mnuMore" label="More">\n'
        '            <toggleButton id="tglX"'
    )
    assert _group(text, "grpMenus") == '<group id="grpMenus" label="Menus">\n        </group>'
    assert text[moved.start : moved.end].endswith("          </menu>")
    # Groups reorder inside their tab the same way.
    regrouped, group = rs.relocate_element(
        RIBBON, _element(RIBBON, "grpMain"), _element(RIBBON, "tabTools"), after=_element(RIBBON, "grpEmpty")
    )
    assert regrouped.index('id="grpEmpty"') < regrouped.index('id="grpMain"')
    assert group.attr("id") == "grpMain"
    _valid(text)
    _valid(regrouped)


def test_relocate_refuses_what_the_schema_forbids():
    with pytest.raises(rs.RibbonEditError, match="cannot hold"):
        rs.relocate_element(RIBBON, _element(RIBBON, "grpMain"), _element(RIBBON, "mnuMore"))
    with pytest.raises(rs.RibbonEditError, match="into itself"):
        rs.relocate_element(RIBBON, _element(RIBBON, "mnuMore"), _element(RIBBON, "mnuMore"))
    tabs = next(e for e in rs.element_tree(RIBBON) if e.name == "tabs")
    with pytest.raises(rs.RibbonEditError, match="fixed position"):
        rs.relocate_element(RIBBON, tabs, _element(RIBBON, "tabTools"))
    assert rs.can_contain(_element(RIBBON, "tabTools"), "group")
    assert not rs.can_contain(_element(RIBBON, "tabTools"), "button")
    assert rs.callback_attrs(_element(RIBBON, "tglX")) == [("onAction", "OnToggle"), ("getPressed", "GetToggle")]
