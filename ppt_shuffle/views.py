from flask import render_template, request, Response, send_file
from io import BytesIO
from pptx import Presentation
import random

from .util import duplicate_slide, delete_slide
from . import app


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/create', methods=['POST'])
def create():
    if 'file' not in request.files:
        return Response(status=304)
    file = request.files['file']
    if file.filename == '':
        return Response(status=304)
    if file:
        fp = BytesIO()
        file.save(fp)
        fp.seek(0)
        prs = Presentation(fp)

        from_ = request.form.get('from')
        if from_.isdigit():
            from_ = int(from_)
        else:
            from_ = 1

        to_ = request.form.get('to')
        if to_.isdigit():
            to_ = int(to_)
        else:
            to_ = len(prs.slides)

        step = request.form.get('step')
        if step.isdigit():
            step = int(step)
        else:
            step = 1

        sub_list = list(zip(*[iter(range(from_, to_))]*step))
        random.shuffle(sub_list)
        sub_list = sum((list(ls) for ls in sub_list), [])

        for i in sub_list:
            duplicate_slide(prs, i)

        for i in range(from_, to_):
            delete_slide(prs, from_)

        f_out = BytesIO()
        prs.save(f_out)
        f_out.seek(0)

        return send_file(f_out, as_attachment=True, attachment_filename=file.filename, cache_timeout=0)

    return Response(status=304)
