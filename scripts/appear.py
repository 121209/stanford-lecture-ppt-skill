#!/usr/bin/env python3
"""Opt-in click entrance animations for lecture delivery.

This writes native PowerPoint timing (``p:timing``) on shapes the caller
already placed. It does not create slides, formulas, diagrams, or video.
Nothing on a slide animates unless this function is called.

``effect="appear"`` is an instant entrance (Office preset 1). ``effect="fade"``
is a 0.5–0.7s fade (Office preset 10). Each shape is its own click unless
``together`` is true, in which case one click reveals every shape. Objects are
not grouped, so OMML stays an editable equation instead of a picture.

This is not ManimGX. Motion clips still go through ``scripts/embed_motion.py``.
"""
# Office entrance presets. Appear is instant; fade is the short dissolve.
_APPEAR_PRESET = 1
_FADE_PRESET = 10
_FADE_MS_MIN = 500
_FADE_MS_MAX = 700
_DEFAULT_FADE_MS = 600
_MAX_SHAPES = 8


class _Ids:
    def __init__(self):
        self._n = 0

    def next(self):
        self._n += 1
        return str(self._n)


def _qn(tag):
    try:
        from pptx.oxml.ns import qn
    except ImportError as exc:
        raise RuntimeError("python-pptx is required for appear.py") from exc
    return qn(tag)


def _add(parent, tag, **attrs):
    element = parent.makeelement(_qn(tag), {key: str(value) for key, value in attrs.items()})
    parent.append(element)
    return element


def _cond(parent, delay, event=None):
    cond = _add(parent, "p:cond", delay=delay)
    if event:
        cond.set("evt", event)
    return cond


def _target(parent, shape_id):
    target = _add(parent, "p:tgtEl")
    _add(target, "p:spTgt", spid=shape_id)
    return target


def _behavior(parent, ids, shape_id, duration):
    behavior = _add(parent, "p:cBhvr")
    node = _add(behavior, "p:cTn", id=ids.next(), dur=str(duration), fill="hold")
    conditions = _add(node, "p:stCondLst")
    _cond(conditions, "0")
    _target(behavior, shape_id)
    return behavior


def _visibility_set(parent, ids, shape_id):
    setter = _add(parent, "p:set")
    behavior = _behavior(setter, ids, shape_id, 1)
    names = _add(behavior, "p:attrNameLst")
    name = _add(names, "p:attrName")
    name.text = "style.visibility"
    _add(_add(setter, "p:to"), "p:strVal", val="visible")
    return setter


def _fade(parent, ids, shape_id, duration_ms):
    effect = _add(parent, "p:animEffect", transition="in", filter="fade")
    _behavior(effect, ids, shape_id, duration_ms)
    return effect


def _effect(parent, ids, shape_id, effect, node_type, duration_ms):
    preset = _APPEAR_PRESET if effect == "appear" else _FADE_PRESET
    wrapper = _add(parent, "p:par")
    node = _add(
        wrapper,
        "p:cTn",
        id=ids.next(),
        presetID=str(preset),
        presetClass="entr",
        presetSubtype="0",
        fill="hold",
        nodeType=node_type,
    )
    _cond(_add(node, "p:stCondLst"), "0")
    children = _add(node, "p:childTnLst")
    if effect == "fade":
        _fade(children, ids, shape_id, duration_ms)
    _visibility_set(children, ids, shape_id)
    return wrapper


def _click_group(parent, ids, shape_ids, effect, duration_ms):
    """One click. The first shape is ``clickEffect``; the rest start with it."""
    outer = _add(parent, "p:par")
    hold = _add(outer, "p:cTn", id=ids.next(), fill="hold")
    _cond(_add(hold, "p:stCondLst"), "indefinite")
    inner = _add(_add(hold, "p:childTnLst"), "p:par")
    inner_hold = _add(inner, "p:cTn", id=ids.next(), fill="hold")
    _cond(_add(inner_hold, "p:stCondLst"), "0")
    effects = _add(inner_hold, "p:childTnLst")
    for index, shape_id in enumerate(shape_ids):
        node_type = "clickEffect" if index == 0 else "withEffect"
        _effect(effects, ids, shape_id, effect, node_type, duration_ms)
    return outer


def _timing(sld, groups, effect, duration_ms):
    ids = _Ids()
    timing = sld.makeelement(_qn("p:timing"), {})
    root = _add(_add(timing, "p:tnLst"), "p:par")
    root_node = _add(
        root,
        "p:cTn",
        id=ids.next(),
        dur="indefinite",
        restart="never",
        nodeType="tmRoot",
    )
    sequence = _add(_add(root_node, "p:childTnLst"), "p:seq", concurrent="1", nextAc="seek")
    main = _add(
        sequence,
        "p:cTn",
        id=ids.next(),
        dur="indefinite",
        fill="hold",
        restart="whenNotActive",
        nodeType="mainSeq",
    )
    children = _add(main, "p:childTnLst")
    for shape_ids in groups:
        _click_group(children, ids, shape_ids, effect, duration_ms)
    for tag, event in (("p:prevCondLst", "onPrev"), ("p:nextCondLst", "onNext")):
        cond = _cond(_add(sequence, tag), "0", event=event)
        _add(_add(cond, "p:tgtEl"), "p:sldTgt")
    return timing


def _shape_ids(slide, shapes):
    if isinstance(shapes, (str, bytes)) or not isinstance(shapes, (list, tuple)):
        raise ValueError("shapes must be a list of shapes on this slide")
    if not shapes:
        raise ValueError("shapes must contain at least one shape")
    if len(shapes) > _MAX_SHAPES:
        raise ValueError("click build accepts at most %d shapes" % _MAX_SHAPES)
    on_slide = {}
    for shape in slide.shapes:
        on_slide[shape._element] = int(shape.shape_id)
    chosen = []
    seen = set()
    for shape in shapes:
        element = getattr(shape, "_element", None)
        if element not in on_slide:
            raise ValueError("each target must be a shape on this slide")
        shape_id = on_slide[element]
        if shape_id in seen:
            raise ValueError("shape %s is listed twice" % shape_id)
        seen.add(shape_id)
        chosen.append(shape_id)
    return chosen


def add_click_appear(slide, shapes, effect="appear", together=False, duration_ms=_DEFAULT_FADE_MS):
    """Add an on-click entrance. Returns the shape ids in click order.

    ``together=False`` reveals one shape per click. ``together=True`` reveals
    the whole list on a single click and does not group the shapes. ``fade``
    lasts ``duration_ms`` milliseconds and must stay within 500–700.
    A slide that already has timing is left unchanged and raises ``ValueError``.
    """
    if effect not in ("appear", "fade"):
        raise ValueError("effect must be appear or fade")
    if not isinstance(together, bool):
        raise ValueError("together must be true or false")
    if effect == "fade":
        if isinstance(duration_ms, bool) or not isinstance(duration_ms, int):
            raise ValueError("duration_ms must be an integer")
        if not _FADE_MS_MIN <= duration_ms <= _FADE_MS_MAX:
            raise ValueError("fade duration must be from 500 to 700 milliseconds")
    else:
        duration_ms = 1
    shape_ids = _shape_ids(slide, shapes)
    sld = slide._element
    if sld.find(_qn("p:timing")) is not None:
        raise ValueError("slide already has timing; refusing to replace it")
    groups = [shape_ids] if together else [[shape_id] for shape_id in shape_ids]
    timing = _timing(sld, groups, effect, duration_ms)
    extension = sld.find(_qn("p:extLst"))
    if extension is not None:
        extension.addprevious(timing)
    else:
        sld.append(timing)
    return shape_ids
