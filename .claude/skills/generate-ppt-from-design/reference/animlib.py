#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""animlib — native PowerPoint animations for python-pptx decks.

python-pptx has NO animation API, so this writes the slide's <p:timing> OOXML
timeline directly. It reproduces the kind of "build" animations Claude designs
use (staggered reveals, fly-ins, scale/pop, color-change via cross-fade, and
continuous pulse loops) as REAL, editable PowerPoint animations — not video.

VALIDATED IN POWERPOINT (opens with no repair; plays correctly): click-sequenced
entrances (fade/fly), and multiple continuous pulse LOOPS that start on slide
load and run through the click-waits. The structural rules that make loops work —
learned the hard way, do not "simplify" them:
  1. A continuous on-load loop must live INSIDE the main sequence (<p:seq nodeType
     ="mainSeq">) as an on-load group (trigger <p:cond delay="0">), NOT as a
     sibling <p:par> of the seq. As a sibling it freezes during click-waits
     (the master clock pauses). Inside the seq (concurrent="1") it keeps running.
  2. ALL on-load/loop effects go in ONE group. One group per loop makes only the
     first auto-start; the rest get bumped onto the first click, pushing every
     click build one click late (off-by-one).
  3. The repeat lives on the BEHAVIOR cTn (inside <p:cBhvr>) as
     repeatCount="indefinite" autoRev="1" — NOT on the preset cTn.
  4. Entrance effects need a <p:bldP spid=.. grpId="0"/> in <p:bldLst> so the
     shape is hidden until its entrance.
Verification: only PowerPoint proves <p:timing> validity/playback — LibreOffice
is lenient and cannot show it. Always hand a tiny test .pptx to the user to run
as a slideshow (see reference/anim_selftest.py), BEFORE animating a full deck.

Model (mirrors PowerPoint's own):
  clicks : list of click-steps. clicks[k] is a list of Effect for the (k+1)-th
           click. The first effect in a step is the click trigger; the rest are
           'with'/'after' relative to it.
  autos  : effects that start automatically on slide load (delay 0), e.g. the
           first build if the design shows it immediately, or continuous loops.

Effect = eff(shape, kind, dir=..., dur=..., delay=..., trig=..., **kw)
  kind: 'fade','appear','fly','zoom','float','wipe'  (entrance)
        'pulse','grow','spin','colorfade'            (emphasis / loop)
        'motion'                                     (motion path, dx/dy in px)
  trig: 'click' (new click), 'with' (with previous), 'after' (after previous)

Usage:
  import animlib as AN
  AN.animate(slide, clicks=[[AN.eff(s1,'fade')],[AN.eff(s2,'fly',dir='u')]],
                    autos=[AN.eff(dot,'pulse',dur=900,loop=True)])
"""
from lxml import etree
from pptx.oxml.ns import qn, nsmap

P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
EMU_PX = 9525

# ---- entrance preset ids (match PowerPoint) ----
ENTR = {
    'appear': (1,  None,        None),
    'fade':   (10, 'in',        'fade'),
    'fly':    (2,  'in',        'fly'),      # uses fromDir subtype
    'float':  (2,  'in',        'fly'),
    'zoom':   (23, 'in',        None),       # grow (animScale)
    'wipe':   (22, 'in',        'wipe'),
}
# directional subtypes for fly/wipe (PowerPoint preset subtypes)
DIRSUB = {'u':4, 'd':1, 'l':2, 'r':8, 'ul':5, 'ur':12, 'dl':3, 'dr':9}
EMPH = {'pulse': 126, 'grow': 6, 'spin': 8, 'colorfade': 10}

def _mk(tag, attrs=None, parent=None):
    ns, local = tag.split(':')
    uri = {'p': P, 'a': A}[ns]
    el = etree.Element('{%s}%s' % (uri, local), nsmap={'p': P, 'a': A})
    if attrs:
        for k, v in attrs.items():
            el.set(k, str(v))
    if parent is not None:
        parent.append(el)
    return el
def _sub(parent, tag, **attrs):
    ns, local = tag.split(':')
    uri = {'p': P, 'a': A}[ns]
    el = etree.SubElement(parent, '{%s}%s' % (uri, local))
    for k, v in attrs.items():
        el.set(k, str(v))
    return el


def eff(shape, kind, dir='u', dur=500, delay=0, trig='click', loop=False,
        dx=0, dy=0, by=160, color=None, repeat=None):
    return dict(shape=shape, kind=kind, dir=dir, dur=int(dur), delay=int(delay),
                trig=trig, loop=loop, dx=dx, dy=dy, by=by, color=color, repeat=repeat)


class _Ids:
    def __init__(self, start=2): self.n = start
    def next(self): self.n += 1; return self.n


def _target(parent, spid):
    tgt = _sub(parent, 'p:tgtEl'); _sub(tgt, 'p:spTgt', spid=str(spid))
    return tgt

def _cbhvr(parent, ids, spid, dur, attrs=None, additive=None):
    cb = _sub(parent, 'p:cBhvr')
    ctn = _sub(cb, 'p:cTn', id=str(ids.next()), dur=str(dur), fill='hold')
    if additive: cb.set('additive', additive)
    _target(cb, spid)
    if attrs:
        al = _sub(cb, 'p:attrNameLst')
        for a in attrs: _sub(al, 'p:attrName').text = a
    return cb

def _effect_nodes(parent, ids, e):
    """append the actual animation behavior nodes for effect e under parent (a clickEffect cTn childTnLst)."""
    spid = e['shape'].shape_id
    kind = e['kind']; dur = e['dur']
    if kind in ('fade', 'appear', 'wipe', 'float', 'fly', 'zoom'):
        # visibility set -> visible (entrance)
        st = _sub(parent, 'p:set')
        cb = _sub(st, 'p:cBhvr')
        _sub(cb, 'p:cTn', id=str(ids.next()), dur='1', fill='hold')
        _target(cb, spid)
        al = _sub(cb, 'p:attrNameLst'); _sub(al, 'p:attrName').text = 'style.visibility'
        to = _sub(st, 'p:to'); _sub(to, 'p:strVal', val='visible')
        if kind == 'fade':
            ae = _sub(parent, 'p:animEffect', transition='in', filter='fade')
            _cbhvr(ae, ids, spid, dur)
        elif kind == 'appear':
            pass
        elif kind == 'wipe':
            ae = _sub(parent, 'p:animEffect', transition='in',
                      filter='wipe(%s)' % {'u':'up','d':'down','l':'left','r':'right'}.get(e['dir'],'up'))
            _cbhvr(ae, ids, spid, dur)
        elif kind in ('fly', 'float'):
            # motion from off-position into place
            am = _sub(parent, 'p:animMotion', origin='layout', pathEditMode='relative',
                      path=_fly_path(e['dir']), rCtr='0,0')
            cb = _cbhvr(am, ids, spid, dur, attrs=['ppt_x', 'ppt_y'])
            ae = _sub(parent, 'p:animEffect', transition='in', filter='fade')
            _cbhvr(ae, ids, spid, min(dur, 200))
        elif kind == 'zoom':
            asx = _sub(parent, 'p:animScale')
            cb = _sub(asx, 'p:cBhvr')
            _sub(cb, 'p:cTn', id=str(ids.next()), dur=str(dur), fill='hold')
            _target(cb, spid)
            _sub(asx, 'p:by', x='160000', y='160000')  # grow from 60%? use from/to
            frm = _sub(asx, 'p:from', x='60000', y='60000');
            ae = _sub(parent, 'p:animEffect', transition='in', filter='fade')
            _cbhvr(ae, ids, spid, min(dur, 200))
    elif kind == 'motion':
        # emphasis motion path: move dx,dy px (relative to slide fractions)
        am = _sub(parent, 'p:animMotion', origin='layout', pathEditMode='relative',
                  path=_rel_path(e['dx'], e['dy']), rCtr='0,0')
        _cbhvr(am, ids, spid, dur, attrs=['ppt_x', 'ppt_y'])
    elif kind in ('pulse', 'grow'):
        asx = _sub(parent, 'p:animScale')
        cb = _sub(asx, 'p:cBhvr')
        cattrs = dict(id=str(ids.next()), dur=str(dur), autoRev='1', fill='hold')
        if e.get('loop') or e.get('repeat'):
            cattrs['repeatCount'] = 'indefinite'   # repeat on the BEHAVIOR cTn so it loops forever
        _sub(cb, 'p:cTn', **cattrs)
        _target(cb, spid)
        b = int(e['by'] * 1000)
        _sub(asx, 'p:by', x=str(b), y=str(b))
    elif kind == 'spin':
        ar = _sub(parent, 'p:animRot', by='21600000')
        _cbhvr(ar, ids, spid, dur, attrs=['r'])
    elif kind == 'colorfade':
        # fade this (duplicate) in on top for a color change via cross-fade
        st = _sub(parent, 'p:set'); cb = _sub(st, 'p:cBhvr')
        _sub(cb, 'p:cTn', id=str(ids.next()), dur='1', fill='hold'); _target(cb, spid)
        al = _sub(cb, 'p:attrNameLst'); _sub(al, 'p:attrName').text = 'style.visibility'
        to = _sub(st, 'p:to'); _sub(to, 'p:strVal', val='visible')
        ae = _sub(parent, 'p:animEffect', transition='in', filter='fade'); _cbhvr(ae, ids, spid, dur)

def _fly_path(d):
    m = {'u':(0,0.5),'d':(0,-0.5),'l':(0.5,0),'r':(-0.5,0),'ul':(0.5,0.5),'ur':(-0.5,0.5),'dl':(0.5,-0.5),'dr':(-0.5,-0.5)}
    dx,dy = m.get(d,(0,0.5))
    return "M %.5f %.5f L 0 0 E" % (dx, dy)
def _rel_path(dx_px, dy_px, sw=1920, sh=1080):
    return "M 0 0 L %.5f %.5f E" % (dx_px/sw, dy_px/sh)


def _preset(e):
    kind = e['kind']
    if kind in ENTR:
        pid, _, _ = ENTR[kind]; cls='entr'
        sub = DIRSUB.get(e['dir'], 0) if kind in ('fly','float','wipe') else 0
        return pid, cls, sub
    if kind in EMPH:
        return EMPH[kind], 'emph', 0
    if kind == 'motion':
        return 0, 'path', 0
    if kind == 'colorfade':
        return 10, 'entr', 0
    return 1, 'entr', 0


def _group_par(parent, ids, step, trig_delay, is_click):
    """One effect group (a <p:par> under mainSeq childTnLst).
    trig_delay='indefinite' -> waits for a click; '0' -> starts on load.
    Auto/on-load groups rely on the mainSeq's concurrent='1' so they keep running
    (and loop) while the sequence waits for the next click."""
    par = _sub(parent, 'p:par')
    ctn = _sub(par, 'p:cTn', id=str(ids.next()), fill='hold')
    scl = _sub(ctn, 'p:stCondLst'); _sub(scl, 'p:cond', delay=str(trig_delay))
    child = _sub(ctn, 'p:childTnLst')
    par2 = _sub(child, 'p:par')
    ctn2 = _sub(par2, 'p:cTn', id=str(ids.next()), fill='hold')
    scl2 = _sub(ctn2, 'p:stCondLst'); _sub(scl2, 'p:cond', delay='0')
    child2 = _sub(ctn2, 'p:childTnLst')
    for i, e in enumerate(step):
        pid, cls, sub = _preset(e)
        if is_click:
            node = 'clickEffect' if i == 0 else ('withEffect' if e['trig'] == 'with' else 'afterEffect')
        else:
            node = 'withEffect'   # on-load group: all start with the timeline
        epar = _sub(child2, 'p:par')
        attrs = dict(id=str(ids.next()), presetID=str(pid), presetClass=cls,
                     presetSubtype=str(sub), fill='hold', grpId='0', nodeType=node)
        ectn = _sub(epar, 'p:cTn', **attrs)
        escl = _sub(ectn, 'p:stCondLst'); _sub(escl, 'p:cond', delay=str(e['delay']))
        echild = _sub(ectn, 'p:childTnLst')
        _effect_nodes(echild, ids, e)


def animate(slide, clicks=None, autos=None):
    """Attach a <p:timing> tree to the slide. clicks: list of steps (each a list
    of eff()). autos: list of eff() that start on load (incl. loops)."""
    clicks = clicks or []
    autos = autos or []
    sld = slide._element  # <p:sld>
    # remove existing timing
    old = sld.find(qn('p:timing'))
    if old is not None: sld.remove(old)
    ids = _Ids(start=2)
    timing = _mk('p:timing')
    tnlst = _sub(timing, 'p:tnLst')
    root_par = _sub(tnlst, 'p:par')
    root_ctn = _sub(root_par, 'p:cTn', id='1', dur='indefinite', restart='never', nodeType='tmRoot')
    root_child = _sub(root_ctn, 'p:childTnLst')
    # main click sequence
    seq = _sub(root_child, 'p:seq', concurrent='1', nextAc='seek')
    seq_ctn = _sub(seq, 'p:cTn', id=str(ids.next()), dur='indefinite', nodeType='mainSeq')
    seq_child = _sub(seq_ctn, 'p:childTnLst')
    # on-load / looping effects FIRST, ALL in ONE group so they start together on
    # load (a seq only auto-starts its first child; extra groups would wait for a
    # click and bump every click build by one).
    if autos:
        _group_par(seq_child, ids, autos, trig_delay='0', is_click=False)
    # then the click groups
    for step in clicks:
        _group_par(seq_child, ids, step, trig_delay='indefinite', is_click=True)
    prev = _sub(seq, 'p:prevCondLst'); c = _sub(prev, 'p:cond', evt='onPrev', delay='0')
    t = _sub(c, 'p:tgtEl'); _sub(t, 'p:sldTgt')
    nxt = _sub(seq, 'p:nextCondLst'); c = _sub(nxt, 'p:cond', evt='onNext', delay='0')
    t = _sub(c, 'p:tgtEl'); _sub(t, 'p:sldTgt')
    # build list — marks each animated shape as a build target (hidden until entrance)
    bld = _sub(timing, 'p:bldLst')
    seen = set()
    for step in clicks:
        for e in step:
            if e['kind'] in ('pulse','grow','spin','motion'):  # emphasis/motion: not a build (shape stays visible)
                continue
            sp = e['shape'].shape_id
            if sp in seen: continue
            seen.add(sp)
            _sub(bld, 'p:bldP', spid=str(sp), grpId='0')
    sld.append(timing)
    return timing
