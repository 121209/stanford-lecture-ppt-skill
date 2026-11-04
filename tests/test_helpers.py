"""Small behavioral tests; fixtures are synthetic package fragments, not deliverable decks."""
import importlib.util
import json
import shutil
import sys
from pathlib import Path
import tempfile
import unittest
import xml.etree.ElementTree as ET
import zipfile
from contextlib import redirect_stdout
from io import StringIO

ROOT = Path(__file__).resolve().parents[1]
def load(name):
    spec = importlib.util.spec_from_file_location(name, ROOT/'scripts'/f'{name}.py')
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module
audit = load('audit_pptx')
state = load('run_state')

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
            ['onscreen_text_over_word_cap'],
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
                raw_textbox(2, '训练流程只有文字'),
                raw_textbox(2, '训练流程') + raw_round(3, '读入'),
                raw_textbox(2, '架构说明') + RAW_CONNECTOR,
                raw_textbox(2, '模型结构') + RAW_PICTURE,
                raw_textbox(2, '算法页缺图：原文图尚未取得。'),
                raw_textbox(2, '下一步讨论收获'),
                raw_textbox(2, '推理流程') + raw_textbox(3, '来源：下半部占位。', y=4500000),
            ])
            report = audit.audit(path)
            self.assertEqual(report['slides_missing_diagram'], [1, 7])
            self.assertGreater(report['slides'][1]['diagram_object_count'], 0)
            self.assertGreater(report['slides'][2]['connector_count'], 0)
            self.assertEqual(report['slides'][3]['picture_count'], 1)
            self.assertTrue(report['slides'][4]['has_gap_marker'])
            self.assertFalse(report['slides'][5]['algorithm_intent'])
            self.assertEqual(
                [item['code'] for item in report['slides'][0]['warnings']],
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
            blob = dest.read_bytes()
            self.assertNotIn(b'{{MATH:', blob)
            self.assertNotIn(b'\\lambda', blob)
            bad = folder/'bad.pptx'
            with self.assertRaises(ValueError):
                inject.inject_pptx(src, {'rate': {'runs': [{'t': r'\lambda'}]}}, bad)
            self.assertFalse(bad.exists())
            pandoc_out = folder/'pandoc.pptx'
            if shutil.which('pandoc'):
                inject.inject_pptx(src, {'rate': {'latex': r'\lambda(t)=\alpha'}}, pandoc_out)
                self.assertEqual(audit.audit(pandoc_out)['slides_with_raw_latex'], [])
            else:
                with self.assertRaises(ValueError) as caught:
                    inject.inject_pptx(src, {'rate': {'latex': r'\lambda(t)=\alpha'}}, pandoc_out)
                self.assertIn('pandoc', str(caught.exception))
                self.assertFalse(pandoc_out.exists())

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

if __name__ == '__main__':unittest.main()
