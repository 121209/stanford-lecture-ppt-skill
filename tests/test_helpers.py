"""Small behavioral tests; fixtures are synthetic package fragments, not deliverable decks."""
import importlib.util
import json
from pathlib import Path
import tempfile
import unittest
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

if __name__ == '__main__':unittest.main()
