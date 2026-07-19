import os
import tempfile
import unittest
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

import fitz
from PIL import Image, ImageCms
from PyQt5.QtCore import QPoint
from PyQt5.QtWidgets import QApplication

from core.pdf_engine import inspect_pdf_color_resources, run_ghostscript
from core.utils import find_ghostscript
from modules.mod_cropper.worker import CropWorker
from modules.mod_pdf_organizer.ui import PDFOrganizerWidget, PageEntry
from modules.mod_stamper.worker import StamperWorker
from modules.mod_toolkit.worker import ToolkitWorker


def _stream_object(dictionary, payload):
    return dictionary.replace(b"__LEN__", str(len(payload)).encode()) + b"\nstream\n" + payload + b"\nendstream"


def _write_color_test_pdf(path):
    icc = ImageCms.ImageCmsProfile(ImageCms.createProfile("sRGB")).tobytes()
    function = b"{ exch 0 exch 0 }"
    content = (
        b"/Spot cs 1 0.5 scn 30 30 160 220 re f\n"
        b"q /GS1 gs /ICC cs 0.2 0.5 0.7 scn 210 30 160 220 re f Q\n"
    )
    objects = [
        b"<< /Type /Catalog /Pages 2 0 R >>",
        b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
        b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 400 300] /Resources << /ColorSpace << /Spot 4 0 R /ICC [/ICCBased 6 0 R] >> /ExtGState << /GS1 << /Type /ExtGState /ca 0.5 /CA 0.5 >> >> >> /Contents 7 0 R >>",
        b"[/DeviceN [/PANTONE#20300#20C /Black] /DeviceCMYK 5 0 R]",
        _stream_object(
            b"<< /FunctionType 4 /Domain [0 1 0 1] /Range [0 1 0 1 0 1 0 1] /Length __LEN__ >>",
            function,
        ),
        _stream_object(b"<< /N 3 /Alternate /DeviceRGB /Length __LEN__ >>", icc),
        _stream_object(b"<< /Length __LEN__ >>", content),
    ]
    data = bytearray(b"%PDF-1.7\n%\xe2\xe3\xcf\xd3\n")
    offsets = []
    for index, obj in enumerate(objects, 1):
        offsets.append(len(data))
        data.extend(f"{index} 0 obj\n".encode() + obj + b"\nendobj\n")
    xref = len(data)
    data.extend(f"xref\n0 {len(objects) + 1}\n".encode() + b"0000000000 65535 f \n")
    for offset in offsets:
        data.extend(f"{offset:010d} 00000 n \n".encode())
    data.extend(
        f"trailer\n<< /Size {len(objects) + 1} /Root 1 0 R >>\nstartxref\n{xref}\n%%EOF\n".encode()
    )
    path.write_bytes(data)


class V4PDFRegressionTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.root = Path(self.temp_dir.name)

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_printer_compression_preserves_devicen_icc_and_spot_name(self):
        gs_path, gs_lib = find_ghostscript()
        if not gs_path:
            self.skipTest("Ghostscript not available")
        source = self.root / "source.pdf"
        output = self.root / "output.pdf"
        _write_color_test_pdf(source)
        run_ghostscript(gs_path, gs_lib, str(source), str(output), "/printer")
        before = inspect_pdf_color_resources(source)
        after = inspect_pdf_color_resources(output)
        self.assertTrue(after["devicen"])
        self.assertTrue(after["icc"])
        self.assertEqual(before["spot_names"], after["spot_names"])

    def test_stamp_reuses_raster_and_keeps_pdf_stamp_text(self):
        target = self.root / "target.pdf"
        raster = self.root / "stamp.png"
        vector = self.root / "stamp.pdf"
        Image.new("RGBA", (2000, 1200), (255, 0, 0, 120)).save(raster)
        doc = fitz.open()
        for _ in range(3):
            doc.new_page()
        doc.save(target)
        doc.close()
        doc = fitz.open()
        page = doc.new_page(width=200, height=80)
        page.insert_text((15, 48), "VECTOR V4", fontsize=22)
        doc.save(vector)
        doc.close()
        positions = {
            index: [
                {"id": "r", "path": str(raster), "w": 25.4, "h": 15.24, "pdf_x": 40,
                 "pdf_y": 50, "angle": 0, "asset_type": "image"},
                {"id": "v", "path": str(vector), "w": 50, "h": 20, "pdf_x": 120,
                 "pdf_y": 100, "angle": 17, "asset_type": "pdf", "pdf_page": 0},
            ] for index in range(3)
        }
        worker = StamperWorker(
            [str(target)], positions, {"mode": "batch", "prefix": "", "suffix": ""},
            {"use_gs": False, "quality": "/ebook"}, str(self.root), False, stamp_dpi=300,
        )
        stamped, _, _ = worker._task_prep_and_stamp(str(target), 0)
        doc = fitz.open(stamped)
        self.assertTrue(all("VECTOR V4" in page.get_text() for page in doc))
        image_xrefs = {image[0] for page in doc for image in page.get_images(full=True)}
        self.assertEqual(len(image_xrefs), 1)
        doc.close()
        Path(stamped).unlink()

    def test_crop_round_trip_keeps_pixel_scale_and_alpha(self):
        source = self.root / "source.png"
        Image.new("RGBA", (600, 400), (10, 100, 220, 140)).save(source, dpi=(300, 300))
        current = source
        for round_index in range(3):
            doc = fitz.open()
            page = doc.new_page(width=600 * 72 / 300, height=400 * 72 / 300)
            page.insert_image(page.rect, filename=str(current), keep_proportion=False)
            output_dir = self.root / f"round_{round_index}"
            output_dir.mkdir()
            errors = []
            worker = CropWorker(
                doc, {0: {"v_lines": [], "h_lines": [], "disabled": []}}, "images",
                str(output_dir), {0: "slice"}, {0: 300},
            )
            worker.error.connect(errors.append)
            worker.run()
            doc.close()
            self.assertFalse(errors)
            current = next(output_dir.glob("*.png"))
            with Image.open(current) as image:
                self.assertEqual(image.size, (600, 400))
                self.assertEqual(image.mode, "RGBA")

    def test_organizer_preserves_order_rotation_links_and_text(self):
        source = self.root / "source.pdf"
        output = self.root / "organized.pdf"
        doc = fitz.open()
        for index in range(3):
            page = doc.new_page()
            page.insert_text((50, 60), f"PAGE {index + 1}")
            page.insert_link({"kind": fitz.LINK_URI, "from": fitz.Rect(50, 80, 160, 100),
                              "uri": "https://example.com"})
        doc.set_toc([[1, "First", 1], [1, "Third", 3]])
        doc.save(source)
        doc.close()
        entries = [PageEntry(str(source), 2, 90), PageEntry(str(source), 0)]
        PDFOrganizerWidget._compose(entries, str(output))
        doc = fitz.open(output)
        self.assertEqual([page.get_text().strip() for page in doc], ["PAGE 3", "PAGE 1"])
        self.assertEqual(doc[0].rotation, 90)
        self.assertTrue(doc[1].get_links())
        self.assertEqual({item[1] for item in doc.get_toc()}, {"First", "Third"})
        doc.close()

    def test_organizer_thumbnail_drag_reorders_pages_and_supports_undo(self):
        class InternalDropEvent:
            def __init__(self, source, position):
                self._source = source
                self._position = position
                self.accepted = False

            def mimeData(self):
                return self

            def hasUrls(self):
                return False

            def source(self):
                return self._source

            def pos(self):
                return self._position

            def setDropAction(self, action):
                self.action = action

            def accept(self):
                self.accepted = True

        app = QApplication.instance() or QApplication([])
        source = self.root / "drag-source.pdf"
        doc = fitz.open()
        for index in range(4):
            doc.new_page().insert_text((40, 50), f"PAGE {index + 1}")
        doc.save(source)
        doc.close()

        widget = PDFOrganizerWidget()
        widget.entries = [PageEntry(str(source), index) for index in range(4)]
        uids = [entry.uid for entry in widget.entries]
        widget.resize(900, 600)
        widget.show()
        widget.refresh_grid()
        app.processEvents()

        target_rect = widget.grid.visualItemRect(widget.grid.item(2))
        widget.grid._dragged_uids = [uids[0]]
        event = InternalDropEvent(
            widget.grid, QPoint(target_rect.right() - 2, target_rect.center().y())
        )
        widget.grid.dropEvent(event)
        self.assertTrue(event.accepted)
        self.assertEqual([entry.source_page for entry in widget.entries], [1, 2, 0, 3])
        self.assertTrue(widget.dirty)

        widget.undo()
        self.assertEqual([entry.source_page for entry in widget.entries], [0, 1, 2, 3])
        widget.grid._dragged_uids = [uids[1], uids[2]]
        event = InternalDropEvent(
            widget.grid, QPoint(widget.grid.viewport().width() - 2, widget.grid.viewport().height() - 2)
        )
        widget.grid.dropEvent(event)
        self.assertEqual([entry.source_page for entry in widget.entries], [0, 3, 1, 2])

        widget.thread_pool.waitForDone(5000)
        widget.close()
        widget.deleteLater()
        app.processEvents()

    def test_interleave_merge_order_reserved_page_and_rotation(self):
        source_a = self.root / "a.pdf"
        source_b = self.root / "b.pdf"
        output = self.root / "interleaved.pdf"
        for path, prefix in [(source_a, "A"), (source_b, "B")]:
            doc = fitz.open()
            for index in range(2):
                page = doc.new_page(width=400, height=600)
                page.insert_text((50, 60), f"{prefix}{index + 1}")
            doc.save(path)
            doc.close()
        worker = ToolkitWorker(
            [str(source_a), str(source_b)], "PDF交叉合并", str(output), 200,
            "jpg", False, interleave_config={
                "rotate_a": 90, "rotate_b": 0, "reverse_b": False,
                "fill_blank": True, "reserved_a": "2", "reserved_b": "",
            },
        )
        worker._run_interleave_merge()
        doc = fitz.open(output)
        self.assertEqual(
            [page.get_text().strip() for page in doc],
            ["A1", "B1", "", "B2", "A2", ""],
        )
        self.assertEqual(doc[0].rotation, 90)
        self.assertEqual(doc[2].rotation, 90)
        doc.close()


if __name__ == "__main__":
    unittest.main()
