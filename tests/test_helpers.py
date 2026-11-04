"""Small behavioral tests; fixtures are synthetic package fragments, not deliverable decks."""
import importlib.util
import json
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

def write_deck(path, slides):
    """slides: list of shape lists. shape keys: text, sz, and optional y (EMU)."""
    ns = audit.NS
    with zipfile.ZipFile(path, 'w') as z:
        for index, shapes in enumerate(slides, 1):
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
            z.writestr(
                f'ppt/slides/slide{index}.xml',
                f'<p:sld xmlns:p="{ns["p"]}" xmlns:a="{ns["a"]}"><p:cSld><p:spTree>{"".join(body)}</p:spTree></p:cSld></p:sld>',
            )
        ids = ''.join(f'<p:sldId id="{255+i}" r:id="r{i}"/>' for i in range(1, len(slides)+1))
        rels = ''.join(
            f'<Relationship Id="r{i}" Type="x/slide" Target="slides/slide{i}.xml"/>'
            for i in range(1, len(slides)+1)
        )
        z.writestr(
            'ppt/presentation.xml',
            f'<p:presentation xmlns:p="{ns["p"]}" xmlns:r="{ns["r"]}"><p:sldIdLst>{ids}</p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/></p:presentation>',
        )
        z.writestr(
            'ppt/_rels/presentation.xml.rels',
            f'<Relationships xmlns="{ns["rel"]}">{rels}</Relationships>',
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
        self.assertEqual(report['slide_count'],6)
        self.assertFalse(report['slides_without_body_notes'])
        self.assertEqual(report['slides'][3]['native_chart_count'],1)
        self.assertTrue(all(s['native_text_shape_count']>2 for s in report['slides']))
        self.assertTrue(report['slides'][0]['onscreen_word_cap_exempt'])
        self.assertEqual(report['slides_over_word_cap'], [3, 5])
        self.assertTrue(all('onscreen_word_count' in s for s in report['slides']))
        self.assertEqual(report['slides_with_empty_lower_content_band'], [])
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

if __name__ == '__main__':unittest.main()
