import sys
from pathlib import Path
from contextlib import contextmanager
import shutil
import tempfile
import os
from zipfile import ZipFile
import hashlib
from datetime import datetime, timedelta
from docx2pdf import convert
import uuid
import random
import time

this_dir = Path(__file__).resolve().parent

docs_dir = this_dir.parent.joinpath("documents")
bin_dir = this_dir.parent.joinpath("bin")
tests_dir = this_dir.parent.joinpath("tests")

os.environ["PATH"] = f"{bin_dir.joinpath('pdftk')};{os.environ['PATH']}"
import pypdftk
from pypdftk import run_command, PDFTK_PATH

temp_cache = Path(tempfile.gettempdir()).joinpath("ignis-pdf-cache")


def init_temp_cache():
    os.makedirs(temp_cache, exist_ok=True)
    keep = set(sorted([i for i in temp_cache.glob("*") if i.is_file()], key=os.path.getmtime)[-500:])
    for i in temp_cache.glob("*"):
        if i.is_file() and i not in keep:
            os.remove(i)


def hexhash(filename):
    with open(filename, "rb") as f:
        bytes = f.read()  # read entire file as bytes
        readable_hash = hashlib.sha256(bytes).hexdigest()
        return(readable_hash)


def safe_convert(wordfile, pdffile=None):
    print(rf"Convert {wordfile}")

    wordfile = Path(wordfile)
    if pdffile is None:
        pdffile = wordfile.parent.joinpath(os.path.splitext(wordfile.name)[0]+".pdf")

    init_temp_cache()
    cached = temp_cache.joinpath(hexhash(wordfile)+".pdf")
    if cached.exists():
        shutil.copy2(cached, pdffile)
        return

    tmpname = f"file-{uuid.uuid4()}"
    with tempfile.TemporaryDirectory() as tdir:
        tdir = Path(tdir)
        tf = tdir.joinpath(tmpname+".docx")
        shutil.copy2(wordfile, tf)
        convert(tf)
        shutil.copy2(tdir.joinpath(tmpname + ".pdf"), pdffile)
        shutil.copy2(tdir.joinpath(tmpname + ".pdf"), cached)


@contextmanager
def chdir(chpath):
    curdir = os.getcwd()
    try:
        os.chdir(chpath)
        yield
    finally:
        os.chdir(curdir)


def stamp_and_replace(input_pdf, stamp_pdf):
    input_pdf = Path(input_pdf)
    with tempfile.TemporaryDirectory() as f:
        shutil.copy(input_pdf, Path(f).joinpath(input_pdf.name))
        pypdftk.stamp(str(Path(f).joinpath(input_pdf.name)),
                      str(stamp_pdf),
                      str(input_pdf))


def background_and_replace(input_pdf, stamp_pdf):
    input_pdf = Path(input_pdf)
    with tempfile.TemporaryDirectory() as f:
        shutil.copy(input_pdf, Path(f).joinpath(input_pdf.name))
        run_command([PDFTK_PATH,
                    str(Path(f).joinpath(input_pdf.name)),
                    "background",
                    str(stamp_pdf),
                    "output",
                    str(input_pdf)])


def word_text_replace(wordfile, replacements, outfile=None):
    r"""
    >>> wordfile = tests_dir.joinpath("worddoc.docx")
    >>> with tempfile.TemporaryDirectory() as tdir:
    ...     _ = shutil.copy2(wordfile, tmpfile:=Path(tdir).joinpath("xxxword.docx"))
    ...     _ = word_text_replace(tmpfile, {"Sdffds":"Hello World", "Sdfsdff":"I'm good" })
    """
    wordfile = Path(wordfile)
    if outfile is None:
        outfile = wordfile

    with tempfile.TemporaryDirectory() as f:
        f = Path(f)
        wordtmp = f.joinpath(wordfile.name)
        shutil.copy2(wordfile, wordtmp)

        shutil.unpack_archive(wordtmp, f.joinpath("zip"), "zip")

        with open(f.joinpath("zip", "word", "document.xml"), "rb") as fr:
            txt_in = fr.read()

        for i, j in replacements.items():
            txt_in = txt_in.replace(i.encode("utf-8"), j.encode("utf-8"))

        with open(f.joinpath("zip", "word", "document.xml"), "wb") as fw:
            fw.write(txt_in)

        shutil.make_archive(f.joinpath("zip"), 'zip', f.joinpath("zip"))
        shutil.copy2(f.joinpath("zip.zip"), outfile)

        return outfile


def target_pdf_hash(pdfin, target, pdfout=None):
    print(f"Forcing SHA256 as {target}... ", end="", flush=True)
    t1 = time.time()

    pdfin = Path(pdfin)
    if pdfout is None:
        pdfout = pdfin
    pdfout = Path(pdfout)

    target = str(target)
    with open(pdfin, "rb") as f:
        txt_in = f.read()

    txt_in = txt_in.rstrip(b"\r\n").rstrip(b"\n")
    origin = hashlib.sha256()
    origin.update(txt_in)
    origin.digest()
    for i in range(1,1000000000000):
        seq = f"\n%>>>IGNIS {i}<<<\n".encode("utf-8")
        a = origin.copy()
        a.update(seq)
        if a.digest().hex().startswith(target):
            break

    with open(pdfout, "wb") as f:
        f.write(txt_in + seq)

    print(f"{time.time()-t1:.2f}s")

def make_documents(input_path, uid, is_draft=False):
    input_path = Path(input_path).resolve()
    input_name = os.path.splitext(input_path.name)[0]
    
    with tempfile.TemporaryDirectory() as f:
        shutil.copytree(docs_dir, f, dirs_exist_ok=True)
        with chdir(f):
            shutil.copy(input_path, "live_document.pdf")

            ignis_id = datetime.now().strftime(f"Ignis Certificate ID {uid} %Y-%m-%d %H:%M")
            word_text_replace("Watermark.docx", {"__id__": ignis_id})
            safe_convert("Watermark.docx")

            stamp_and_replace("live_document.pdf", "Watermark.pdf")

            pypdftk.get_pages("live_document.pdf",
                              ranges=[[1], [2]],
                              out_file="live_document_1_2.pdf")

            with tempfile.TemporaryDirectory() as fout:
                fout = Path(fout)
                if is_draft:
                    word_text_replace("DRAFT.docx",
                                      {'__date__': (datetime.now()+timedelta(days=14)).strftime("%Y-%m-%d")})

                    safe_convert("DRAFT.docx")
                    stamp_and_replace("live_document_1_2.pdf", "DRAFT.pdf")
                    stamp_and_replace("live_document.pdf", "DRAFT.pdf")

                fn1 = fout.joinpath(input_name+" Certificate.pdf")
                fn2 = fout.joinpath(input_name+" Report.pdf")
                fn3 = fout.joinpath(input_name+" Verification.pdf")

                target_pdf_hash("live_document_1_2.pdf", uid)
                target_pdf_hash("live_document.pdf", uid)

                shutil.copy2("live_document_1_2.pdf", fn1)
                shutil.copy2("live_document.pdf", fn2)

                if not is_draft:
                    word_text_replace("FileReport.docx",
                                      {"__filename_1__": fn1.name,
                                       "__sha256_1__": hexhash(fn1),
                                       "__filename_2__": fn2.name,
                                       "__sha256_2__": hexhash(fn2)})

                    safe_convert("FileReport.docx")
                    stamp_and_replace("FileReport.pdf", "Watermark.pdf")
                    target_pdf_hash("FileReport.pdf", uid)
                    shutil.copy2("FileReport.pdf", fn3)

                with chdir(fout):
                    files = list(Path(".").glob("*"))
                    # create a ZipFile object

                    zipObj = ZipFile(f'{input_name}.zip', 'w')
                    for i in files:
                        zipObj.write(i)
                    zipObj.close()

                    shutil.copy(f'{input_name}.zip', input_path.parent.joinpath(f'{input_name}.zip'))
                    return input_path.parent.joinpath(f'{input_name}.zip')


if __name__ == "__main__":
    fname = Path(sys.argv[1]).resolve()

    name, ext = os.path.splitext(fname.name)
    ext = ext.lower()

    assert(ext in (".pdf", ".docx"))

    is_draft = False
    if os.path.splitext(fname.name)[0].lower().endswith("draft"):
        is_draft = True

    uid = uuid.uuid4().hex[:5]

    if ext == ".docx":
        with tempfile.TemporaryDirectory() as fdir:
            fdocx = Path(fdir).joinpath(fname.name)
            fpdf = Path(fdir).joinpath(name+".pdf")
            shutil.copy2(fname, fdocx)

            safe_convert(fdocx, fpdf)
            outzip = make_documents(fpdf, uid, is_draft=is_draft)
            shutil.copy2(outzip, fname.parent.joinpath(outzip.name))
    else:
        make_documents(sys.argv[1], uid, is_draft=is_draft)
