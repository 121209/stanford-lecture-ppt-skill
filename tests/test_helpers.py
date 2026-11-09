"""Small behavioral tests; fixtures are synthetic package fragments, not deliverable decks."""
import importlib.util
import json
import re
import shutil
import struct
import subprocess
import sys
from pathlib import Path
import tempfile
import unittest
import xml.etree.ElementTree as ET
import zipfile
from contextlib import redirect_stdout
from io import BytesIO, StringIO

ROOT = Path(__file__).resolve().parents[1]
def load(name):
    spec = importlib.util.spec_from_file_location(name, ROOT/'scripts'/f'{name}.py')
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module
audit = load('audit_pptx')
state = load('run_state')

def _ooxml_schema_dir():
    folder = ROOT / 'tests' / 'xsd'
    if (folder / 'pml.xsd').is_file():
        return folder
    return None

def write_raw_deck(path, bodies):
    """bodies: list of shape-tree inner XML strings, one per slide."""
    ns = audit.NS
    with zipfile.ZipFile(path, 'w') as z:
        for index, body in enumerate(bodies, 1):
            z.writestr(
                f'ppt/slides/slide{index}.xml',
                (
                    f'<p:sld xmlns:p="{ns["p"]}" xmlns:a="{ns["a"]}" '
                    f'xmlns:r="{ns["r"]}" xmlns:m="{ns["m"]}">'
                    f'<p:cSld><p:spTree>{body}</p:spTree></p:cSld></p:sld>'
                ),
            )
        count = len(bodies)
        ids = ''.join(f'<p:sldId id="{255+i}" r:id="r{i}"/>' for i in range(1, count+1))
        rels = ''.join(
            f'<Relationship Id="r{i}" Type="x/slide" Target="slides/slide{i}.xml"/>'
            for i in range(1, count+1)
        )
        z.writestr(
            'ppt/presentation.xml',
            f'<p:presentation xmlns:p="{ns["p"]}" xmlns:r="{ns["r"]}"><p:sldIdLst>{ids}</p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/></p:presentation>',
        )
        z.writestr(
            'ppt/_rels/presentation.xml.rels',
            f'<Relationships xmlns="{ns["rel"]}">{rels}</Relationships>',
        )

def write_deck(path, slides):
    """slides: list of shape lists. shape keys: text, sz, and optional y (EMU)."""
    bodies = []
    for shapes in slides:
        body = []
        for number, shape in enumerate(shapes, 2):
            y = shape.get('y', 1000000)
            text = shape['text']
            sz = shape['sz']
            body.append(
                '<p:sp>'
                f'<p:nvSpPr><p:cNvPr id="{number}" name="s"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
                '<p:spPr><a:xfrm>'
                f'<a:off x="700000" y="{y}"/><a:ext cx="10000000" cy="400000"/>'
                '</a:xfrm></p:spPr>'
                '<p:txBody><a:bodyPr/><a:lstStyle/>'
                f'<a:p><a:r><a:rPr sz="{sz}"/><a:t>{text}</a:t></a:r></a:p>'
                '</p:txBody></p:sp>'
            )
        bodies.append(''.join(body))
    write_raw_deck(path, bodies)

def raw_textbox(number, text, y=1600000, sz=2200):
    return (
        '<p:sp>'
        f'<p:nvSpPr><p:cNvPr id="{number}" name="t"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
        '<p:spPr><a:xfrm>'
        f'<a:off x="700000" y="{y}"/><a:ext cx="8000000" cy="500000"/>'
        '</a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
        f'<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="{sz}"/><a:t>{text}</a:t></a:r></a:p></p:txBody>'
        '</p:sp>'
    )

def raw_round(number, text):
    return (
        '<p:sp>'
        f'<p:nvSpPr><p:cNvPr id="{number}" name="n"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
        '<p:spPr><a:xfrm><a:off x="700000" y="1800000"/><a:ext cx="2200000" cy="800000"/></a:xfrm>'
        '<a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>'
        '<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>'
        '<a:ln><a:solidFill><a:srgbClr val="C00000"/></a:solidFill></a:ln></p:spPr>'
        f'<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="2000"/><a:t>{text}</a:t></a:r></a:p></p:txBody>'
        '</p:sp>'
    )

RAW_CONNECTOR = (
    '<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="9" name="c"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>'
    '<p:spPr><a:xfrm><a:off x="3000000" y="2100000"/><a:ext cx="400000" cy="0"/></a:xfrm>'
    '<a:prstGeom prst="line"><a:avLst/></a:prstGeom></p:spPr></p:cxnSp>'
)
RAW_PICTURE = (
    '<p:pic><p:nvPicPr><p:cNvPr id="10" name="p"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>'
    '<p:blipFill><a:blip r:embed="rId9"/></p:blipFill>'
    '<p:spPr><a:xfrm><a:off x="700000" y="1800000"/><a:ext cx="4000000" cy="1800000"/></a:xfrm>'
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>'
)

class HelperTests(unittest.TestCase):
    def test_invalid_zip_rejected(self):
        with tempfile.TemporaryDirectory() as folder:
            p=Path(folder)/'broken.pptx';p.write_bytes(b'not-a-zip')
            with self.assertRaises(zipfile.BadZipFile):audit.audit(p)

    def test_actual_presentation_order(self):
        with tempfile.TemporaryDirectory() as folder:
            p=Path(folder)/'fixture.pptx'
            ns=audit.NS
            with zipfile.ZipFile(p,'w') as z:
                z.writestr('ppt/presentation.xml',f'<p:presentation xmlns:p="{ns["p"]}" xmlns:r="{ns["r"]}"><p:sldIdLst><p:sldId id="256" r:id="B"/><p:sldId id="257" r:id="A"/></p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/></p:presentation>')
                z.writestr('ppt/_rels/presentation.xml.rels',f'<Relationships xmlns="{ns["rel"]}"><Relationship Id="A" Type="x/slide" Target="slides/slide1.xml"/><Relationship Id="B" Type="x/slide" Target="slides/slide9.xml"/></Relationships>')
                for n in [1,9]:z.writestr(f'ppt/slides/slide{n}.xml',f'<p:sld xmlns:p="{ns["p"]}"><p:cSld><p:spTree/></p:cSld></p:sld>')
            original=p.read_bytes();report=audit.audit(p)
            self.assertEqual([s['part'] for s in report['slides']],['ppt/slides/slide9.xml','ppt/slides/slide1.xml'])
            self.assertEqual(report['slides_without_body_notes'],[1,2])
            self.assertEqual(p.read_bytes(),original)

    def test_gates_notes_and_changed_evidence(self):
        with tempfile.TemporaryDirectory() as folder:
            project=Path(folder)
            def call(command,*args):
                with redirect_stdout(StringIO()):state.main([command,'--project',str(project),*map(str,args)])
            evidence=project/'outline.md';evidence.write_text('test outline')
            artifact=project/'fragment.json';artifact.write_text('{}')
            preview=project/'preview.png';preview.write_bytes(b'fixture only, not rendering')
            deck=project/'fixture.pptx'
            with zipfile.ZipFile(deck,'w') as z:z.writestr('ppt/presentation.xml',f'<p:presentation xmlns:p="{audit.NS["p"]}"><p:sldIdLst><p:sldId id="256"/></p:sldIdLst></p:presentation>')
            call('init','--pages',1)
            with self.assertRaises(ValueError):call('plan','--spec',evidence)
            for gate in ['outline','style','backend']:call('gate','--name',gate,'--basis','test approval','--evidence',evidence)
            call('gate','--name','sample','--waived','--basis','explicit test waiver')
            spec=project/'spec.json';spec.write_text(json.dumps({'route':'editable','method':{'tool':'test','mode':'native'},'slides':[{'id':'slide_01','order':1,'title':'Title','role':'concept','key_points':['Point'],'layout_intent':'Explain','local_context':'Definition','notes_outline':'Talk','source_assets':[]}]}))
            call('plan','--spec',spec)
            call('start','--slide','slide_01','--worker','main')
            call('record','--slide','slide_01','--artifact',artifact,'--preview',preview,'--qa-note','fixture')
            with self.assertRaises(ValueError):call('verify','--deck',deck)
            call('start','--slide','slide_01','--worker','main')
            call('record','--slide','slide_01','--artifact',artifact,'--preview',preview,'--qa-note','fixture','--notes-ready')
            call('verify','--deck',deck)
            preview.write_bytes(b'changed')
            with self.assertRaises(ValueError):call('verify','--deck',deck)

    def test_demo_native_objects(self):
        report=audit.audit(ROOT/'assets/demo.pptx')
        self.assertEqual(report['slide_count'],7)
        self.assertFalse(report['slides_without_body_notes'])
        self.assertEqual(report['slides'][3]['native_chart_count'],1)
        self.assertTrue(all(s['native_text_shape_count']>2 for s in report['slides']))
        self.assertTrue(report['slides'][0]['onscreen_word_cap_exempt'])
        self.assertEqual(report['slides_over_word_cap'], [3, 5])
        self.assertTrue(all('onscreen_word_count' in s for s in report['slides']))
        self.assertEqual(report['slides_with_empty_lower_content_band'], [])
        self.assertEqual(report['slides_missing_diagram'], [])
        self.assertEqual(report['slides_with_raw_latex'], [])
        self.assertFalse(any(s['algorithm_intent'] for s in report['slides'][:6]))
        self.assertTrue(report['slides'][6]['algorithm_intent'])
        self.assertGreater(report['slides'][6]['diagram_object_count'], 0)
        self.assertGreater(report['slides'][6]['connector_count'], 0)
        self.assertIn('does not change the process exit code', report['claim_boundary'])

    def test_word_count_cjk_ratio(self):
        self.assertEqual(audit.word_count('中'*80), 40)
        self.assertEqual(audit.word_count('中'*81), 40.5)
        self.assertEqual(audit.word_count('mean squared error'), 3)
        self.assertEqual(audit.word_count('误差 MSE'), 2)

    def test_word_cap_flags_overflow_not_cover_or_source_lines(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder)/'words.pptx'
            footer_y = 6200000
            write_deck(path, [
                [{'text': '中'*90, 'sz': 2200}],
                [
                    {'text': '中'*80, 'sz': 2200},
                    {'text': '02', 'sz': 1300, 'y': footer_y},
                    {'text': '来源：示意数据，不代表论文结果。'+'补'*40, 'sz': 1400, 'y': 1400000},
                    {'text': '页底说明不计入屏幕字数。'+'注'*30, 'sz': 1500, 'y': footer_y},
                    {'text': '来源：下半部占位，不计入屏幕字数。', 'sz': 1400, 'y': 4500000},
                ],
                [{'text': '中'*81, 'sz': 2200}],
                [
                    {'text': '中'*90, 'sz': 1400, 'y': 1600000},
                    {'text': '来源：下半部占位，不计入屏幕字数。', 'sz': 1400, 'y': 4500000},
                ],
            ])
            report = audit.audit(path)
        self.assertEqual([s['onscreen_word_count'] for s in report['slides']], [45, 40, 40.5, 45])
        self.assertEqual(report['slides_over_word_cap'], [3, 4])
        self.assertTrue(report['slides'][0]['onscreen_word_cap_exempt'])
        self.assertEqual(
            [item['code'] for item in report['slides'][2]['warnings']],
            ['onscreen_text_over_word_cap'],
        )
        self.assertEqual(
            [item['code'] for item in report['slides'][3]['warnings']],
            ['onscreen_text_over_word_cap', 'font_below_min'],
        )
        self.assertFalse(report['slides'][1]['warnings'])
        self.assertEqual(report['slides_with_empty_lower_content_band'], [])

    def test_empty_lower_content_band_warns_without_changing_exit_code(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder)/'band.pptx'
            upper = [
                {'text': '读入', 'sz': 2200, 'y': 1600000},
                {'text': '编码', 'sz': 2200, 'y': 2100000},
                {'text': '更新', 'sz': 2200, 'y': 2600000},
                {'text': '评价', 'sz': 2200, 'y': 3100000},
            ]
            write_deck(path, [
                list(upper),
                list(upper),
                list(upper) + [{'text': '结论：顺序依赖数据，不代表原文算法。', 'sz': 2000, 'y': 4500000}],
                list(upper) + [{'text': '来源：示意', 'sz': 1200, 'y': 6200000}],
                list(upper) + [{'text': '格子略过中线', 'sz': 2200, 'y': 3400000}],
            ])
            report = audit.audit(path)
            self.assertEqual(report['slides_with_empty_lower_content_band'], [2, 4, 5])
            self.assertFalse(any(
                item['code'] == 'empty_lower_content_band' for item in report['slides'][0]['warnings']
            ))
            self.assertFalse(any(
                item['code'] == 'empty_lower_content_band' for item in report['slides'][2]['warnings']
            ))
            self.assertEqual(
                [item['code'] for item in report['slides'][1]['warnings']],
                ['empty_lower_content_band'],
            )
            saved = sys.argv
            try:
                sys.argv = ['audit_pptx.py', str(path)]
                with redirect_stdout(StringIO()):
                    code = audit.main()
            finally:
                sys.argv = saved
            self.assertEqual(code, 0)

    def test_group_box_fills_lower_band_and_slide_number_does_not(self):
        ns = audit.NS
        tree = ET.fromstring(
            f'<p:spTree xmlns:p="{ns["p"]}" xmlns:a="{ns["a"]}">'
            '<p:sp><p:nvSpPr><p:cNvPr id="6" name="empty"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
            '<p:spPr><a:xfrm><a:off x="700000" y="4500000"/><a:ext cx="8000000" cy="800000"/></a:xfrm></p:spPr>'
            '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="2200"/><a:t> </a:t></a:r></a:p></p:txBody>'
            '</p:sp>'
            '<p:sp><p:nvSpPr><p:cNvPr id="2" name="t"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
            '<p:spPr><a:xfrm><a:off x="700000" y="1600000"/><a:ext cx="4000000" cy="800000"/></a:xfrm></p:spPr>'
            '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="2200"/><a:t>上</a:t></a:r></a:p></p:txBody>'
            '</p:sp>'
            '<p:sp><p:nvSpPr><p:cNvPr id="5" name="num"/><p:cNvSpPr/><p:nvPr><p:ph type="sldNum"/></p:nvPr></p:nvSpPr>'
            '<p:spPr><a:xfrm><a:off x="700000" y="4500000"/><a:ext cx="4000000" cy="800000"/></a:xfrm></p:spPr>'
            '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1200"/><a:t>2</a:t></a:r></a:p></p:txBody>'
            '</p:sp>'
            '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="3" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
            '<p:grpSpPr><a:xfrm><a:off x="700000" y="4500000"/><a:ext cx="8000000" cy="800000"/>'
            '<a:chOff x="0" y="0"/><a:chExt cx="8000000" cy="800000"/></a:xfrm></p:grpSpPr>'
            '<p:sp><p:nvSpPr><p:cNvPr id="4" name="c"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
            '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="400000"/></a:xfrm></p:spPr>'
            '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1800"/><a:t>示意</a:t></a:r></a:p></p:txBody>'
            '</p:sp></p:grpSp></p:spTree>'
        )
        boxes = audit.content_band_boxes(tree)
        self.assertEqual(len(boxes), 2)
        self.assertEqual(boxes[1][2], 8000000)
        self.assertGreaterEqual(boxes[1][1], 4000000)
        self.assertFalse(audit.lower_content_band_is_empty(boxes, 6858000))
        self.assertTrue(audit.lower_content_band_is_empty(boxes[:1], 6858000))

    def test_algorithm_page_needs_diagram_gap_or_picture(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder)/'algo.pptx'
            write_raw_deck(path, [
                raw_textbox(2, '封面'),
                raw_textbox(2, '训练流程只有文字') + raw_textbox(4, '下半部有一句说明。', y=4500000),
                raw_textbox(2, '训练流程') + raw_round(3, '读入'),
                raw_textbox(2, '架构说明') + RAW_CONNECTOR,
                raw_textbox(2, '模型结构') + RAW_PICTURE,
                raw_textbox(2, '算法页缺图：原文图尚未取得。'),
                raw_textbox(2, '下一步讨论收获'),
                raw_textbox(2, '推理流程') + raw_textbox(3, '来源：下半部占位。', y=4500000),
            ])
            report = audit.audit(path)
            self.assertEqual(report['slides_missing_diagram'], [2, 8])
            self.assertGreater(report['slides'][2]['diagram_object_count'], 0)
            self.assertGreater(report['slides'][3]['connector_count'], 0)
            self.assertEqual(report['slides'][4]['picture_count'], 1)
            self.assertTrue(report['slides'][5]['has_gap_marker'])
            self.assertFalse(report['slides'][6]['algorithm_intent'])
            self.assertEqual(
                [item['code'] for item in report['slides'][1]['warnings']],
                ['algorithm_page_missing_diagram'],
            )
            saved = sys.argv
            try:
                sys.argv = ['audit_pptx.py', str(path)]
                with redirect_stdout(StringIO()):
                    code = audit.main()
            finally:
                sys.argv = saved
            self.assertEqual(code, 0)

    def test_raw_latex_warns_and_omml_is_counted(self):
        omml = (
            '<p:sp><p:nvSpPr><p:cNvPr id="4" name="m"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
            '<p:spPr><a:xfrm><a:off x="700000" y="1800000"/><a:ext cx="5000000" cy="600000"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>'
            '<p:txBody><a:bodyPr/><a:lstStyle/><a:p>'
            '<m:oMath><m:r><m:t>λ</m:t></m:r></m:oMath>'
            '</a:p></p:txBody></p:sp>'
        )
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder)/'math.pptx'
            write_raw_deck(path, [
                raw_textbox(2, 'λ(t)=α，只说明符号'),
                raw_textbox(2, r'损失 \frac{1}{2} 与 x^{2}')
                + raw_textbox(3, '下半部有一句说明。', y=4500000),
                omml + raw_textbox(5, '下半部有一句说明。', y=4500000),
            ])
            report = audit.audit(path)
        self.assertEqual(report['slides_with_raw_latex'], [2])
        latex_warning = report['slides'][1]['warnings'][0]
        self.assertEqual(latex_warning['code'], 'raw_latex_on_screen')
        self.assertIn(r'\frac', latex_warning['examples'])
        self.assertIn('^{', latex_warning['examples'])
        self.assertEqual(report['slides'][0]['warnings'], [])
        self.assertEqual(report['office_math_count'], 1)
        self.assertEqual(report['slides_with_office_math'], [3])
        self.assertIn(
            'office_math_missing_fallback',
            [item['code'] for item in report['slides'][2]['warnings']],
        )

    def test_inject_omml_replaces_placeholder_and_rejects_source(self):
        inject = load('inject_omml')
        with tempfile.TemporaryDirectory() as folder:
            folder = Path(folder)
            src = folder/'in.pptx'
            write_deck(src, [[{'text': '{{MATH:rate}}', 'sz': 2200, 'y': 1800000}]])
            dest = folder/'out.pptx'
            inject.inject_pptx(src, {
                'rate': {'runs': [{'t': 'λ'}, {'t': '='}, {'sup': {'base': 'e', 'sup': '−t'}}]},
            }, dest)
            report = audit.audit(dest)
            self.assertEqual(report['office_math_count'], 1)
            self.assertEqual(report['slides_with_raw_latex'], [])
            self.assertEqual(report['slides_with_invalid_office_math'], [])
            blob = dest.read_bytes()
            self.assertNotIn(b'{{MATH:', blob)
            self.assertNotIn(b'\\lambda', blob)
            self.assertIn(b'mc:AlternateContent', blob)
            self.assertIn(b'Requires="a14"', blob)
            self.assertIn(b'<a14:m', blob)
            self.assertIn(b'm:oMathPara', blob)
            self.assertIn(b'm:oMath', blob)
            self.assertIn(b'mc:Fallback', blob)
            self.assertIn(b'a:blip', blob)
            bad = folder/'bad.pptx'
            with self.assertRaises(ValueError):
                inject.inject_pptx(src, {'rate': {'runs': [{'t': r'\lambda'}]}}, bad)
            self.assertFalse(bad.exists())
            latex_out = folder/'latex.pptx'
            inject.inject_pptx(src, {'rate': {'latex': r'\lambda(t)=\alpha'}}, latex_out)
            latex_report = audit.audit(latex_out)
            self.assertEqual(latex_report['office_math_count'], 1)
            self.assertEqual(latex_report['slides_with_raw_latex'], [])
            self.assertNotIn(b'\\lambda', latex_out.read_bytes())
            self.assertNotIn(b'{{MATH:', latex_out.read_bytes())

    def test_diagram_helper_builds_audited_flow(self):
        try:
            import pptx  # noqa: F401
        except ImportError:
            self.skipTest('python-pptx is not installed')
        helpers = load('diagram_helpers')
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder)/'flow.pptx'
            helpers.build_slide_from_spec({
                'title': '训练流程：四步示意',
                'layout': 'row',
                'nodes': helpers.FLOW_NODES,
                'edges': helpers.FLOW_EDGES,
                'conclusion': '顺序是汇报归纳，不是原文算法。',
                'source': '来源：匿名示意，不是论文实验。',
                'page': '02',
            }, path)
            report = audit.audit(path)
        self.assertEqual(report['slides_missing_diagram'], [])
        self.assertEqual(report['slides_with_raw_latex'], [])
        self.assertEqual(report['slides_over_word_cap'], [])
        self.assertEqual(report['slides_with_empty_lower_content_band'], [])
        self.assertGreater(report['slides'][0]['connector_count'], 0)
        self.assertTrue(report['slides'][0]['algorithm_intent'])
        with self.assertRaises(ValueError):
            helpers._check_graph(
                [{'id': 'n%d' % i, 'label': '点'} for i in range(13)],
                [],
            )

    def test_checked_in_diagram_samples(self):
        path = ROOT/'assets/diagram-samples/figure-first.pptx'
        self.assertTrue(path.is_file(), 'missing anonymous diagram samples')
        report = audit.audit(path)
        self.assertEqual(report['slide_count'], 3)
        self.assertEqual(report['slides_missing_diagram'], [])
        self.assertEqual(report['slides_with_raw_latex'], [])
        self.assertEqual(report['slides_over_word_cap'], [])
        self.assertEqual(report['slides_with_empty_lower_content_band'], [])
        self.assertTrue(report['slides'][0]['algorithm_intent'])
        self.assertTrue(report['slides'][1]['algorithm_intent'])
        self.assertGreater(report['slides'][0]['connector_count'], 0)
        self.assertGreater(report['slides'][1]['diagram_object_count'], 0)
        self.assertGreater(report['slides'][2]['picture_count'], 0)
        self.assertFalse(report['slides'][2]['algorithm_intent'])
        self.assertFalse(report['slides_without_body_notes'])

    def test_missing_diagram_matches_title_not_model_name(self):
        """Cover, table, chart, and a summary that only mentions a model must not warn."""
        ns = audit.NS
        title = lambda number, text: raw_textbox(number, text, y=400000, sz=3200)
        bullet = lambda number, text: raw_textbox(number, text, y=1800000)
        lower = raw_textbox(9, '下半部有一句说明。', y=4500000)
        table = (
            '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="20" name="tbl"/>'
            '<p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>'
            '<p:xfrm><a:off x="700000" y="1600000"/><a:ext cx="8000000" cy="2200000"/></p:xfrm>'
            '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">'
            '<a:tbl><a:tblGrid><a:gridCol w="8000000"/></a:tblGrid>'
            '<a:tr h="400000"><a:tc><a:txBody><a:bodyPr/><a:lstStyle/>'
            '<a:p><a:r><a:rPr sz="2000"/><a:t>Transformer++</a:t></a:r></a:p>'
            '</a:txBody></a:tc></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>'
        )
        chart = (
            '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="21" name="chart"/>'
            '<p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>'
            '<p:xfrm><a:off x="700000" y="1600000"/><a:ext cx="8000000" cy="2200000"/></p:xfrm>'
            '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">'
            f'<c:chart xmlns:c="{ns["c"]}" r:id="rIdC"/></a:graphicData></a:graphic></p:graphicFrame>'
        )
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder)/'titles.pptx'
            write_raw_deck(path, [
                title(2, 'Mamba：方法与架构') + lower,
                title(2, '小结') + bullet(3, '这里回顾 Transformer，也提到架构二字。') + lower,
                title(2, '方法：Transformer++ 结果') + table + lower,
                title(2, '方法：整体架构') + chart + lower,
                title(2, '方法：整体架构') + bullet(3, '只有文字，没有图。') + lower,
            ])
            report = audit.audit(path)
        self.assertEqual(report['slides_missing_diagram'], [5])
        self.assertIn('Mamba', report['slides'][0]['title_text'])
        self.assertTrue(report['slides'][0]['algorithm_intent'])
        self.assertNotIn(
            'algorithm_page_missing_diagram',
            [item['code'] for item in report['slides'][0]['warnings']],
        )
        self.assertEqual(report['slides'][1]['title_text'], '小结')
        self.assertFalse(report['slides'][1]['algorithm_intent'])
        self.assertGreater(report['slides'][2]['native_table_count'], 0)
        self.assertGreater(report['slides'][3]['native_chart_count'], 0)
        self.assertEqual(
            [item['code'] for item in report['slides'][4]['warnings']
             if item['code'] == 'algorithm_page_missing_diagram'],
            ['algorithm_page_missing_diagram'],
        )

    def test_native_omml_opens_and_libreoffice_shows_fallback(self):
        try:
            from pptx import Presentation
            from pptx.util import Emu, Inches, Pt
        except ImportError:
            self.skipTest('python-pptx is not installed')
        inject = load('inject_omml')
        soffice = shutil.which('soffice') or shutil.which('libreoffice')
        if not soffice:
            self.skipTest('LibreOffice is not installed')
        with tempfile.TemporaryDirectory() as folder:
            folder = Path(folder)
            prs = Presentation()
            prs.slide_width = Emu(12192000)
            prs.slide_height = Emu(6858000)
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            box = slide.shapes.add_textbox(Inches(0.8), Inches(2.1), Inches(8.2), Inches(1.05))
            box.text_frame.paragraphs[0].text = '{{MATH:rate}}'
            box.text_frame.paragraphs[0].runs[0].font.size = Pt(28)
            src = folder/'in.pptx'
            prs.save(str(src))
            dest = folder/'out.pptx'
            inject.inject_pptx(src, {
                'rate': {'runs': [
                    {'t': 'λ('}, {'t': 't'}, {'t': ')='}, {'t': 'α'},
                    {'sup': {'base': 'e', 'sup': '−βt'}},
                ]},
            }, dest)
            reloaded = Presentation(str(dest))
            self.assertEqual(len(reloaded.slides), 1)
            self._assert_native_omml_package(dest)
            report = audit.audit(dest)
            self.assertEqual(report['office_math_count'], 1)
            self.assertEqual(report['slides_with_invalid_office_math'], [])
            self.assertEqual(report['slides_with_raw_latex'], [])
            ratio = self._render_nonwhite_ratio(soffice, dest, folder/'render')
        self.assertGreater(ratio, 0.001, 'fallback equation image was not visible')

    def _assert_native_omml_package(self, path):
        ns = {
            'a': audit.NS['a'],
            'm': audit.NS['m'],
            'mc': audit.NS['mc'],
            'a14': audit.NS['a14'],
            'p': audit.NS['p'],
            'r': audit.NS['r'],
            'rel': audit.NS['rel'],
            'ct': 'http://schemas.openxmlformats.org/package/2006/content-types',
        }
        with zipfile.ZipFile(path) as package:
            slide = ET.fromstring(package.read('ppt/slides/slide1.xml'))
            rels = ET.fromstring(package.read('ppt/slides/_rels/slide1.xml.rels'))
            types = package.read('[Content_Types].xml').decode('utf-8')
            names = set(package.namelist())
        self.assertIn('Extension="png"', types)
        blocks = slide.findall('.//mc:AlternateContent', ns)
        self.assertTrue(blocks, 'missing mc:AlternateContent')
        choice = blocks[0].find('mc:Choice', ns)
        fallback = blocks[0].find('mc:Fallback', ns)
        self.assertIsNotNone(choice)
        self.assertIsNotNone(fallback)
        self.assertIn('a14', (choice.get('Requires') or '').split())
        math = choice.find('.//a14:m/m:oMathPara/m:oMath', ns)
        self.assertIsNotNone(math, 'Choice is not a14:m/m:oMathPara/m:oMath')
        blip = fallback.find('.//a:blip', ns)
        self.assertIsNotNone(blip, 'Fallback has no picture')
        embed = blip.get('{%s}embed' % ns['r'])
        self.assertTrue(embed)
        target = None
        for rel in rels:
            if rel.get('Id') == embed:
                target = rel.get('Target')
        self.assertTrue(target)
        media = 'ppt/media/' + target.split('/')[-1]
        self.assertIn(media, names)
        self._assert_fallback_choice_xsd(path)

    def _assert_fallback_choice_xsd(self, path):
        """Validate Choice and Fallback against the transitional slide schema when present.

        Markup-compatibility preprocessing matches an Office validator: the fallback
        branch is what a reader without a14 checks, and a14 elements are dropped
        before the base schema sees the choice branch.
        """
        try:
            from lxml import etree
        except ImportError:
            self.skipTest('lxml is not installed')
        schema_dir = _ooxml_schema_dir()
        if schema_dir is None:
            self.fail('OOXML schemas could not be loaded for slide validation')
        schema = etree.XMLSchema(etree.parse(str(schema_dir / 'pml.xsd')))
        mc = audit.NS['mc']
        a14 = audit.NS['a14']
        with zipfile.ZipFile(path) as package:
            slide = etree.fromstring(package.read('ppt/slides/slide1.xml'))
        for keep in ('Fallback', 'Choice'):
            doc = etree.fromstring(etree.tostring(slide))
            for node in list(doc.xpath('.//mc:AlternateContent', namespaces={'mc': mc})):
                picked = node.find('{%s}%s' % (mc, keep))
                parent = node.getparent()
                index = list(parent).index(node)
                parent.remove(node)
                if picked is not None:
                    for offset, child in enumerate(list(picked)):
                        parent.insert(index + offset, child)
            if '{%s}Ignorable' % mc in doc.attrib:
                del doc.attrib['{%s}Ignorable' % mc]
            for node in list(doc.iter()):
                if not isinstance(node.tag, str):
                    continue
                if node.tag.startswith('{%s}' % a14):
                    node.getparent().remove(node)
            if not schema.validate(doc):
                errors = [err.message for err in schema.error_log][:4]
                self.fail('slide failed XSD after selecting %s: %s' % (keep, '; '.join(errors)))

    def _render_nonwhite_ratio(self, soffice, pptx, dest):
        from PIL import Image
        dest.mkdir(parents=True, exist_ok=True)
        profile = dest / 'profile'
        cmd = [
            soffice, '--headless', '--norestore', '--nolockcheck', '--nologo',
            '-env:UserInstallation=file://%s' % profile,
            '--convert-to', 'pdf', '--outdir', str(dest), str(pptx),
        ]
        completed = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
        self.assertEqual(completed.returncode, 0, completed.stderr[-400:])
        pdf = next(dest.glob('*.pdf'))
        subprocess.run(['pdftoppm', '-png', '-r', '72', str(pdf), str(dest / 'slide')], check=True)
        image = Image.open(next(dest.glob('slide*.png'))).convert('RGB')
        pixels = image.load()
        width, height = image.size
        dark = 0
        for y in range(height):
            for x in range(width):
                red, green, blue = pixels[x, y]
                if red < 240 or green < 240 or blue < 240:
                    dark += 1
        return dark / float(width * height)

    def test_diamond_label_stays_one_line_and_column_arrows_grow(self):
        try:
            import pptx  # noqa: F401
        except ImportError:
            self.skipTest('python-pptx is not installed')
        helpers = load('diagram_helpers')
        with tempfile.TemporaryDirectory() as folder:
            folder = Path(folder)
            diamond = folder/'diamond.pptx'
            helpers.build_slide_from_spec({
                'title': '判断：是否完成',
                'layout': 'row',
                'region': [helpers.emu(0.7), helpers.emu(2.0), helpers.emu(2.2), helpers.emu(1.2)],
                'nodes': [{'id': 'done', 'label': '完成？', 'shape': 'diamond'}],
                'edges': [],
                'conclusion': '判断写在菱形里，不换行。',
                'source': '来源：匿名示意，不是论文实验。',
                'page': '02',
            }, diamond)
            with zipfile.ZipFile(diamond) as package:
                xml = package.read('ppt/slides/slide1.xml').decode('utf-8')
            self.assertIn('wrap="none"', xml)
            self.assertNotIn('<a:br', xml)
            size = ET.fromstring(xml).find('.//a:rPr', {'a': audit.NS['a']})
            # The title run is 32pt; the diamond run is the smaller one.
            sizes = [
                int(node.get('sz'))
                for node in ET.fromstring(xml).findall('.//a:rPr', {'a': audit.NS['a']})
                if node.get('sz')
            ]
            self.assertTrue(sizes)
            self.assertLessEqual(min(sizes), 2000)
            columns = folder/'columns.pptx'
            helpers.build_slide_from_spec({
                'title': '架构：编码端与解码端',
                'layout': 'columns',
                'nodes': helpers.ARCH_NODES,
                'edges': helpers.ARCH_EDGES,
                'conclusion': '模块为示意，不代表论文原图。',
                'source': '来源：匿名结构，不是论文 Figure。',
                'page': '02',
            }, columns)
            root = ET.fromstring(zipfile.ZipFile(columns).read('ppt/slides/slide1.xml'))
            ns = {'a': audit.NS['a'], 'p': audit.NS['p']}
            vertical = []
            for connector in root.findall('.//p:cxnSp', ns):
                ext = connector.find('.//a:ext', ns)
                if ext is None:
                    continue
                cx, cy = int(ext.get('cx', '0')), int(ext.get('cy', '0'))
                if cy > cx:
                    vertical.append(cy)
                    tail = connector.find('.//a:tailEnd', ns)
                    self.assertIsNotNone(tail)
                    self.assertEqual(tail.get('len'), 'lg')
            self.assertGreaterEqual(len(vertical), 2)
            self.assertGreaterEqual(min(vertical), helpers.emu(0.36))
            report = audit.audit(columns)
        self.assertEqual(report['slides_over_word_cap'], [])
        self.assertEqual(report['slides_with_empty_lower_content_band'], [])
        self.assertEqual(report['slides_missing_diagram'], [])

    def test_vendor_latex_compiles_hawkes_and_attention(self):
        inject = load('inject_omml')
        hawkes = {
            'rate': r'\lambda(t)=\mu+\sum_{t_i<t}\alpha e^{-\beta(t-t_i)}',
            'll': r'\log L=\sum_{i=1}^{n}\log\lambda(t_i)-\int_0^T\lambda(t)\,dt',
        }
        attention = r'\mathrm{softmax}\left(\frac{QK^\top}{\sqrt{d_k}}\right)V'
        rate = inject.omml_for_spec({'latex': hawkes['rate']})
        ll = inject.omml_for_spec({'latex': hawkes['ll']})
        attn = inject.omml_for_spec({'latex': attention})
        for omml in (rate, ll, attn):
            self.assertIn('<m:oMath', omml)
            self.assertNotIn('<m:oMathPara', omml)
            self.assertIsNone(inject.RAW_LATEX.search(omml))
        self.assertIn('undOvr', rate)
        self.assertIn('undOvr', ll)
        self.assertIn('<m:nary', rate)
        self.assertIn('<m:nary', ll)
        self.assertIn('subSup', ll)
        self.assertIn('<m:f', attn)
        self.assertIn('<m:rad', attn)
        both = inject.omml_for_spec({'runs': [{'t': 'λ'}], 'latex': r'\alpha'})
        self.assertIn('λ', both)
        self.assertNotIn('>α<', both)
        drawn = inject.mathtext_for_spec({'runs': [{'t': 'log'}], 'latex': r'\log L'}, '')
        self.assertIn(r'\log', drawn)
        self.assertNotIn('displaystyle', inject.mathtext_for_spec({'latex': r'\displaystyle\sum_i i'}, ''))
        with self.assertRaises(ValueError) as caught:
            inject.omml_for_spec({'latex': r'\foo'})
        self.assertIn('Unsupported', str(caught.exception))
        self.assertIn(r'\foo', str(caught.exception))

    def test_latex_omml_validates_and_fallback_keeps_aspect(self):
        try:
            from pptx import Presentation
            from pptx.util import Emu, Inches, Pt
        except ImportError:
            self.skipTest('python-pptx is not installed')
        inject = load('inject_omml')
        soffice = shutil.which('soffice') or shutil.which('libreoffice')
        with tempfile.TemporaryDirectory() as folder:
            folder = Path(folder)
            prs = Presentation()
            prs.slide_width = Emu(12192000)
            prs.slide_height = Emu(6858000)
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            # Wide and short on purpose: stretching this box would change the glyph aspect.
            box = slide.shapes.add_textbox(Inches(0.55), Inches(2.4), Inches(11.2), Inches(0.55))
            box.text_frame.paragraphs[0].text = '{{MATH:ll}}'
            box.text_frame.paragraphs[0].runs[0].font.size = Pt(22)
            src = folder/'in.pptx'
            prs.save(str(src))
            dest = folder/'out.pptx'
            inject.inject_pptx(src, {
                'll': {'latex': r'\log L=\sum_{i=1}^{n}\log\lambda(t_i)-\int_0^T\lambda(t)\,dt'},
            }, dest)
            blob = dest.read_bytes()
            self.assertNotIn(b'{{MATH:', blob)
            self.assertNotIn(b'\\log', blob)
            self.assertNotIn(b'\\sum', blob)
            self.assertIn(b'<a:stretch><a:fillRect/></a:stretch>', blob)
            self.assertIn(b'undOvr', blob)
            report = audit.audit(dest)
            self.assertEqual(report['office_math_count'], 1)
            self.assertEqual(report['slides_with_invalid_office_math'], [])
            self.assertEqual(report['slides_with_raw_latex'], [])
            self._assert_native_omml_package(dest)
            self._assert_fallback_aspect(dest)
            runs_png = inject.render_equation_png(
                [inject.mathtext_for_spec({
                    'runs': [
                        {'t': 'log L='},
                        {'sub': {'base': 'Σ', 'sub': 'i=1…n'}},
                        {'t': ' log λ(ti) − ∫ λ(t) dt'},
                    ],
                }, '')],
                8.0, 1.2, 'log', strict=False,
            )
            latex_png = inject.render_equation_png(
                [inject.mathtext_for_spec({
                    'latex': r'\log L=\sum_{i=1}^{n}\log\lambda(t_i)-\int_0^T\lambda(t)\,dt',
                }, '')],
                8.0, 1.2, 'log', strict=True,
            )
            runs_h = self._ink_height(runs_png)
            latex_h = self._ink_height(latex_png)
            self.assertGreater(latex_h / float(runs_h), 1.3)
            aligned = folder/'aligned.pptx'
            # The OMML compiler accepts aligned. mathtext does not, so without
            # TeX the inject must still fail closed. A real TeX install may draw it.
            tex = inject._load_tex_fallback()
            real_status = tex.toolchain_status
            tex.toolchain_status = lambda: {
                'engine': None, 'engine_path': None,
                'png_converter': None, 'svg_converter': None,
                'tex_png': False, 'tex_svg': False,
            }
            try:
                with self.assertRaises(ValueError):
                    inject.inject_pptx(src, {
                        'll': {'latex': r'\begin{aligned}a&=b\\c&=d\end{aligned}'},
                    }, aligned)
            finally:
                tex.toolchain_status = real_status
            self.assertFalse(aligned.exists())
            if soffice:
                ratio = self._render_nonwhite_ratio(soffice, dest, folder/'render')
                self.assertGreater(ratio, 0.001, 'LibreOffice formula fallback was blank')

    def _assert_fallback_aspect(self, path):
        ns = {
            'a': audit.NS['a'],
            'mc': audit.NS['mc'],
            'r': audit.NS['r'],
        }
        with zipfile.ZipFile(path) as package:
            root = ET.fromstring(package.read('ppt/slides/slide1.xml'))
            fallback = root.find('.//mc:Fallback', ns)
            self.assertIsNotNone(fallback)
            blip = fallback.find('.//a:blip', ns)
            ext = fallback.find('.//a:ext', ns)
            self.assertIsNotNone(blip)
            self.assertIsNotNone(ext)
            embed = blip.get('{%s}embed' % audit.NS['r'])
            rels = ET.fromstring(package.read('ppt/slides/_rels/slide1.xml.rels'))
            target = next(rel.get('Target') for rel in rels if rel.get('Id') == embed)
            png = package.read('ppt/media/' + target.split('/')[-1])
        png_w, png_h = struct.unpack('>II', png[16:24])
        cx, cy = int(ext.get('cx')), int(ext.get('cy'))
        png_aspect = png_w / float(png_h)
        box_aspect = cx / float(cy)
        self.assertLess(
            abs(png_aspect - box_aspect) / box_aspect, 0.02,
            'png %sx%s vs xfrm %sx%s' % (png_w, png_h, cx, cy),
        )

    def _ink_height(self, blob):
        from PIL import Image
        image = Image.open(BytesIO(blob)).convert('RGBA')
        pixels = image.load()
        width, height = image.size
        top, bottom = height, 0
        for y in range(height):
            for x in range(width):
                red, green, blue, alpha = pixels[x, y]
                if alpha < 16:
                    continue
                if red < 240 or green < 240 or blue < 240:
                    top = min(top, y)
                    bottom = max(bottom, y)
        self.assertLess(top, bottom)
        return bottom - top + 1

    def test_method_title_with_valid_omml_is_not_missing_diagram(self):
        try:
            from pptx import Presentation
            from pptx.util import Emu, Inches, Pt
        except ImportError:
            self.skipTest('python-pptx is not installed')
        inject = load('inject_omml')
        with tempfile.TemporaryDirectory() as folder:
            folder = Path(folder)
            prs = Presentation()
            prs.slide_width = Emu(12192000)
            prs.slide_height = Emu(6858000)
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            title = slide.shapes.add_textbox(Inches(0.7), Inches(0.38), Inches(11), Inches(0.6))
            title.text_frame.paragraphs[0].text = '方法：λ(t)'
            title.text_frame.paragraphs[0].runs[0].font.size = Pt(32)
            box = slide.shapes.add_textbox(Inches(0.7), Inches(1.8), Inches(8.5), Inches(1.15))
            box.text_frame.paragraphs[0].text = '{{MATH:rate}}'
            lower = slide.shapes.add_textbox(Inches(0.7), Inches(5.2), Inches(10), Inches(0.5))
            lower.text_frame.paragraphs[0].text = '条件强度，符号见备注。'
            src = folder/'in.pptx'
            prs.save(str(src))
            dest = folder/'out.pptx'
            inject.inject_pptx(src, {
                'rate': {'latex': r'\lambda(t)=\mu+\sum_{t_i<t}\alpha e^{-\beta(t-t_i)}'},
            }, dest)
            report = audit.audit(dest)
            self.assertGreater(report['office_math_count'], 0)
            self.assertEqual(report['slides_with_invalid_office_math'], [])
            self.assertEqual(report['slides'][0]['picture_count'], 0)
            self.assertEqual(report['slides'][0]['diagram_object_count'], 0)
            self.assertEqual(report['slides'][0]['native_chart_count'], 0)
            self.assertTrue(report['slides'][0]['algorithm_intent'])
            self.assertNotIn(
                'algorithm_page_missing_diagram',
                [item['code'] for item in report['slides'][0]['warnings']],
            )
            plain = folder/'plain.pptx'
            write_raw_deck(plain, [
                raw_textbox(2, '封面'),
                raw_textbox(2, '方法：λ(t)', y=400000, sz=3200)
                + raw_textbox(3, '只有文字，没有公式。', y=1800000),
            ])
            plain_report = audit.audit(plain)
            self.assertEqual(plain_report['slides_missing_diagram'], [2])

    def test_bar_chart_is_schema_legal(self):
        try:
            import pptx  # noqa: F401
            from lxml import etree
        except ImportError:
            self.skipTest('python-pptx or lxml is not installed')
        charts = load('chart_helpers')
        helpers = load('diagram_helpers')
        soffice = shutil.which('soffice') or shutil.which('libreoffice')
        with tempfile.TemporaryDirectory() as folder:
            folder = Path(folder)
            prs = helpers.new_deck()
            slide = helpers.add_blank_slide(prs)
            # Selective Copying accuracies reported in Gu & Dao 2023, Table 1.
            helpers.add_bar_chart(
                slide,
                ['S4（LTI SSM）', 'S6（Mamba 选择性）'],
                [18.3, 99.8],
                (helpers.emu(0.8), helpers.emu(1.6), helpers.emu(7.6), helpers.emu(3.6)),
                highlight=1,
                source='来源：Gu & Dao 2023, Table 1（原文报告，本人未复现）',
            )
            path = folder/'chart.pptx'
            prs.save(str(path))
            report = audit.audit(path)
            self.assertEqual(report['slides'][0]['native_chart_count'], 1)
            with zipfile.ZipFile(path) as package:
                chart_xml = package.read('ppt/charts/chart1.xml')
                slide_xml = package.read('ppt/slides/slide1.xml')
            text = chart_xml.decode('utf-8')
            self.assertIn('>18.3<', text)
            self.assertIn('>99.8<', text)
            self.assertNotIn('<c:legend', text)
            self.assertNotIn('<c:title', text)
            self.assertIn('C00000', text)
            self.assertIn('BFBFBF', text)
            self.assertIn('<c:min val="0"', text.replace('0.0', '0'))
            self.assertNotRegex(text, r'<c:(?:axId|crossAx) val="-\d+"')
            for value in re.findall(r'<c:(?:axId|crossAx) val="(-?\d+)"', text):
                self.assertGreaterEqual(int(value), 0)
            schema_dir = _ooxml_schema_dir()
            self.assertIsNotNone(schema_dir)
            chart_schema = etree.XMLSchema(etree.parse(str(schema_dir / 'dml-chart.xsd')))
            slide_schema = etree.XMLSchema(etree.parse(str(schema_dir / 'pml.xsd')))
            self.assertTrue(chart_schema.validate(etree.fromstring(chart_xml)), chart_schema.error_log)
            self.assertTrue(slide_schema.validate(etree.fromstring(slide_xml)), slide_schema.error_log)
            saved = sys.argv
            try:
                sys.argv = ['audit_pptx.py', str(path)]
                with redirect_stdout(StringIO()):
                    code = audit.main()
            finally:
                sys.argv = saved
            self.assertEqual(code, 0)
            if soffice:
                ratio = self._render_nonwhite_ratio(soffice, path, folder/'render')
                self.assertGreater(ratio, 0.001, 'LibreOffice chart was blank')
        with self.assertRaises(ValueError):
            charts._numbers([-1, 2])

    def test_font_floor_and_contrast_warn_without_changing_exit_code(self):
        footer_y = 6200000
        def shape(number, text, sz, y, color=None):
            fill = ''
            if color:
                fill = '<a:solidFill><a:srgbClr val="%s"/></a:solidFill>' % color
            return (
                '<p:sp>'
                '<p:nvSpPr><p:cNvPr id="%d" name="t"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
                '<p:spPr><a:xfrm><a:off x="700000" y="%d"/><a:ext cx="8000000" cy="400000"/></a:xfrm>'
                '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>'
                '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="%d">%s</a:rPr>'
                '<a:t>%s</a:t></a:r></a:p></p:txBody></p:sp>'
            ) % (number, y, sz, fill, text)
        page_number = (
            '<p:sp><p:nvSpPr><p:cNvPr id="8" name="num"/><p:cNvSpPr/><p:nvPr>'
            '<p:ph type="sldNum"/></p:nvPr></p:nvSpPr>'
            '<p:spPr><a:xfrm><a:off x="11000000" y="1600000"/><a:ext cx="500000" cy="300000"/></a:xfrm></p:spPr>'
            '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1200"/>'
            '<a:t>2</a:t></a:r></a:p></p:txBody></p:sp>'
        )
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder)/'type.pptx'
            write_raw_deck(path, [
                shape(2, '封面标题', 3200, 1600000, '222222'),
                ''.join([
                    shape(2, '看不清的正文', 1400, 1600000, '222222'),
                    shape(3, '十八点正文', 1800, 2200000, '222222'),
                    shape(4, '十六点也低于正文硬底', 1600, 2800000),
                    shape(5, '来源：正文区来源行', 1400, 3400000, '606060'),
                    shape(6, '来源：页脚来源', 1300, footer_y, '606060'),
                    page_number,
                ]),
                ''.join([
                    shape(2, '禁用浅灰', 2000, 1600000, 'BFBFBF'),
                    shape(3, '白字', 2000, 2200000, 'FFFFFF'),
                    shape(4, '墨色', 2000, 2800000, '222222'),
                    shape(5, '辅助灰', 2000, 3400000, '606060'),
                    shape(6, '红色标题', 2000, 4000000, 'C00000'),
                ]),
            ])
            report = audit.audit(path)
            saved = sys.argv
            try:
                sys.argv = ['audit_pptx.py', str(path)]
                with redirect_stdout(StringIO()):
                    code = audit.main()
            finally:
                sys.argv = saved
        self.assertEqual(code, 0)
        self.assertEqual(report['slides_with_font_below_min'], [2])
        self.assertEqual(report['slides_with_low_contrast_text'], [3])
        self.assertEqual(report['slides'][0]['warnings'], [])
        font = next(item for item in report['slides'][1]['warnings'] if item['code'] == 'font_below_min')
        self.assertEqual(font['floor_pt'], 18)
        self.assertEqual(font['min_pt'], 14)
        self.assertEqual(font['count'], 2)
        pale = next(item for item in report['slides'][2]['warnings'] if item['code'] == 'low_contrast_text')
        self.assertEqual(pale['examples'], ['BFBFBF', 'FFFFFF'])
        self.assertEqual(pale['minimum_ratio'], 4.5)
        self.assertGreater(audit.contrast_on_white('222222'), 4.5)
        self.assertGreater(audit.contrast_on_white('C00000'), 4.5)
        self.assertGreater(audit.contrast_on_white('606060'), 4.5)
        self.assertLess(audit.contrast_on_white('BFBFBF'), 4.5)

    def test_dfrac_fallback_keeps_omml_frac_and_png_is_transparent(self):
        from PIL import Image
        inject = load('inject_omml')
        latex = r'\frac{QK^{\top}}{\sqrt{d_k}}'
        seen = []
        original = inject.omml_from_vendor
        def spy(source):
            seen.append(source)
            return original(source)
        inject.omml_from_vendor = spy
        try:
            omml = inject.omml_for_spec({'latex': latex})
        finally:
            inject.omml_from_vendor = original
        self.assertEqual(seen, [latex])
        self.assertIn('<m:f>', omml)
        self.assertNotIn('dfrac', omml)
        drawn = inject.mathtext_for_spec({'latex': latex}, omml)
        self.assertIn(r'\dfrac', drawn)
        self.assertNotRegex(drawn, r'(?<![A-Za-z])\\frac\b')
        runs = inject.mathtext_for_spec(
            {'runs': [{'frac': {'num': '1', 'den': 'n'}}]}, '',
        )
        self.assertIn(r'\dfrac{1}{n}', runs)
        self.assertNotIn(r'\frac{', runs)
        self.assertGreaterEqual(inject._fallback_fontsize(0.45, 1, drawn), 22)
        self.assertGreater(
            inject._fallback_fontsize(1.0, 1, drawn),
            inject._fallback_fontsize(1.0, 1, 'x'),
        )
        self.assertGreaterEqual(
            inject._fallback_fontsize(0.8, 1, 'd_k', inline=True),
            inject._fallback_fontsize(0.8, 1, 'd_k', inline=False),
        )
        png = inject.render_equation_png([drawn], 6.0, 1.2, 'frac', strict=True)
        self.assertEqual(png[25], 6)
        image = Image.open(BytesIO(png)).convert('RGBA')
        pixels = image.load()
        width, height = image.size
        self.assertEqual(pixels[0, 0][3], 0)
        self.assertEqual(pixels[width - 1, 0][3], 0)
        self.assertEqual(pixels[0, height - 1][3], 0)
        near_white = 0
        dark = 0
        for y in range(height):
            for x in range(width):
                red, green, blue, alpha = pixels[x, y]
                if alpha < 16:
                    continue
                if red >= 248 and green >= 248 and blue >= 248:
                    near_white += 1
                elif red < 80 and green < 80 and blue < 80:
                    dark += 1
        self.assertEqual(near_white, 0)
        self.assertGreater(dark, 20)

    def test_short_inline_latex_is_omml_not_ascii(self):
        inline = load('inline_math')
        inject = load('inject_omml')
        softmax = inline.prepare_inline_spec(r'\mathrm{softmax}')
        self.assertEqual(softmax['latex'], r'\mathrm{softmax}')
        self.assertTrue(softmax['inline'])
        upright = inject.omml_for_spec(softmax)
        self.assertIn('softmax', upright)
        self.assertNotIn('mathrm', upright)
        with self.assertRaises(ValueError):
            inline.prepare_inline_spec('查询维度 d_k 的缩放')
        with self.assertRaises(ValueError):
            inline.prepare_inline_spec(
                r'\mathrm{softmax}\left(\frac{QK^\top}{\sqrt{d_k}}\right)V'
            )
        with tempfile.TemporaryDirectory() as folder:
            folder = Path(folder)
            src = folder/'in.pptx'
            write_deck(src, [[
                {'text': '{{MATH:dk}}', 'sz': 2000, 'y': 1800000},
                {'text': '{{MATH:lam}}', 'sz': 2000, 'y': 2500000},
                {'text': '{{MATH:sm}}', 'sz': 2000, 'y': 3200000},
            ]])
            # Keep the lower band occupied so this fixture is only about the math.
            dest = folder/'out.pptx'
            inline.inject_inline_pptx(src, {
                'dk': 'd_k',
                'lam': r'\lambda(t)',
                'sm': {'latex': r'\mathrm{softmax}'},
            }, dest)
            report = audit.audit(dest)
            self.assertEqual(report['office_math_count'], 3)
            self.assertEqual(report['slides_with_raw_latex'], [])
            self.assertEqual(report['slides_with_invalid_office_math'], [])
            blob = dest.read_bytes()
            self.assertNotIn(b'{{MATH:', blob)
            self.assertNotIn(b'\\lambda', blob)
            self.assertNotIn(b'\\mathrm', blob)
            with zipfile.ZipFile(dest) as package:
                slide = package.read('ppt/slides/slide1.xml').decode('utf-8')
                pngs = [
                    package.read(name)
                    for name in package.namelist()
                    if name.startswith('ppt/media/') and name.endswith('.png')
                ]
            self.assertIn('<m:sSub>', slide)
            self.assertIn('λ', slide)
            self.assertIn('softmax', slide)
            self.assertNotIn('d_k', slide)
            self.assertNotIn('x_i', slide)
            self.assertNotRegex(slide, r'[A-Za-z]_[A-Za-z0-9]')
            self.assertIn('mc:AlternateContent', slide)
            self.assertIn('mc:Fallback', slide)
            self.assertTrue(pngs)
            for png in pngs:
                self.assertEqual(png[25], 6)

    def test_presentation_style_uses_lab_colors(self):
        try:
            import matplotlib
            matplotlib.use('Agg', force=False)
            import matplotlib.pyplot as plt
        except ImportError:
            self.skipTest('matplotlib is not installed')
        path = ROOT/'assets'/'presentation.mplstyle'
        previous = matplotlib.rcParams.copy()
        try:
            plt.style.use(str(path))
            cycle = [item['color'] for item in plt.rcParams['axes.prop_cycle']]
            self.assertEqual(cycle, ['#222222', '#C00000', '#606060', '#BFBFBF'])
            self.assertFalse(plt.rcParams['axes.spines.top'])
            self.assertFalse(plt.rcParams['axes.spines.right'])
        finally:
            matplotlib.rcParams.update(previous)

    def test_audit_samples_exit_zero(self):
        for sample in (
            ROOT/'assets/demo.pptx',
            ROOT/'assets/diagram-samples/figure-first.pptx',
        ):
            saved = sys.argv
            try:
                sys.argv = ['audit_pptx.py', str(sample)]
                with redirect_stdout(StringIO()):
                    code = audit.main()
            finally:
                sys.argv = saved
            self.assertEqual(code, 0, sample.name)

if __name__ == '__main__':unittest.main()
