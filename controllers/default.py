import datetime
from datetime import datetime as dt


def index():
    ustudents = db().select(
        db.undergrad_stud.student_id,
        db.undergrad_stud.first_name,
        db.undergrad_stud.middle_name,
        db.undergrad_stud.last_name,
        db.undergrad_stud.gender,
        db.undergrad_stud.date_graduated,
        db.degree_courses.course_name,
        db.majors.major,
        left=[db.degree_courses.on(db.undergrad_stud.course_id ==
                                   db.degree_courses.course_id),
              db.majors.on(db.undergrad_stud.major_id == db.majors.major_id)], orderby=db.undergrad_stud.last_name)

    gstudents = db().select(
        db.grad_stud.student_id,
        db.grad_stud.first_name,
        db.grad_stud.middle_name,
        db.grad_stud.last_name,
        db.grad_stud.gender,
        db.degree_courses.course_name,
        db.majors.major,
        db.campus_college_institute.cci_name,
        db.campus_college_institute.address,
        db.campus_college_institute.tel_no,
        left=[db.degree_courses.on(db.grad_stud.course_id ==
                                   db.degree_courses.course_id),
              db.majors.on(db.grad_stud.major_id == db.majors.major_id),
              db.campus_college_institute.on(db.degree_courses.cci_id ==
                                             db.campus_college_institute.cci_id)])

    return locals()


def suffix(d):
    return 'th' if 11 <= d <= 13 else {1: 'st', 2: 'nd', 3: 'rd'}.get(d % 10,
                                                                      'th')


def custom_strftime(format_date, t):
    return t.strftime(format_date).replace('{S}', str(t.day) + suffix(t.day))


def page1():
    # get student id from URL
    try:
        stud_arg = int(request.args(0))
    except TypeError:
        redirect(URL('index.html'))

    try:
        student = db(db.grad_stud.student_id == stud_arg).select(
            db.grad_stud.first_name,
            db.grad_stud.middle_name,
            db.grad_stud.last_name,
            db.grad_stud.gender,
            db.degree_courses.course_name,
            db.degree_courses.course_abbrev,
            db.majors.major,
            db.campus_college_institute.cci_name,
            db.campus_college_institute.address,
            db.campus_college_institute.tel_no,
            left=[db.degree_courses.on(db.grad_stud.course_id ==
                                       db.degree_courses.course_id),
                  db.majors.on(db.grad_stud.major_id == db.majors.major_id),
                  db.campus_college_institute.on(db.degree_courses.cci_id ==
                                                 db.campus_college_institute.cci_id)])

        current_day = custom_strftime('{S}', dt.now())
        current_month_year = custom_strftime('%B, %Y', dt.now())

        if student[0].majors.major is not None:
            major = student[0].majors.major
        else:
            major = "N/A"

        if student[0].grad_stud.gender == 'F':
            address = 'Ms.'
        else:
            address = 'Mr.'

        if student[0].grad_stud.middle_name is not None:
            full_name = address + ' ' + student[0].grad_stud.first_name.upper() + ' ' + \
                student[0].grad_stud.middle_name[0].upper() + '. ' + \
                student[0].grad_stud.last_name.upper()
        else:
            full_name = address + ' ' + student[0].grad_stud.first_name.upper() + ' ' + \
                student[0].grad_stud.last_name.upper()

        address_stud = address + ' ' + student[0].grad_stud.last_name

        course = student[0].degree_courses.course_name + ' (' + \
            student[0].degree_courses.course_abbrev.upper() + ')'

        registrar = db((db.registrar.registrar_position == 'Registrar II')).select(
            db.registrar.registrar_name)

        uni_registrar = db((db.registrar.registrar_position == 'University Registrar')).select(
            db.registrar.registrar_name)

        return locals()
    except IndexError:
        pass


def page2():
    # get student id from URL
    try:
        stud_arg = int(request.args(0))
    except TypeError:
        redirect(URL('index.html'))

    try:
        current_day = custom_strftime('{S}', dt.now())
        current_month_year = custom_strftime('%B, %Y', dt.now())

        student = db(db.undergrad_stud.student_id == stud_arg).select(
            db.undergrad_stud.first_name,
            db.undergrad_stud.middle_name,
            db.undergrad_stud.last_name,
            db.undergrad_stud.gender,
            db.undergrad_stud.date_graduated,
            db.degree_courses.course_name,
            db.degree_courses.course_abbrev,
            db.majors.major,
            left=[db.degree_courses.on(db.undergrad_stud.course_id ==
                                       db.degree_courses.course_id),
                  db.majors.on(db.undergrad_stud.major_id == db.majors.major_id)])

        if student[0].undergrad_stud.date_graduated is not None:
            date_graduated = custom_strftime('%B %d, %Y', student[
                0].undergrad_stud.date_graduated)
            year_graduated = custom_strftime('%Y', student[
                0].undergrad_stud.date_graduated)
        else:
            date_graduated = "N/A"
            year_graduated = "N/A"

        if student[0].undergrad_stud.gender == 'F':
            address = 'Ms.'
        else:
            address = 'Mr.'

        if student[0].undergrad_stud.middle_name is not None:
            full_name = address + ' ' + student[0].undergrad_stud.first_name.upper() + ' ' + \
                student[0].undergrad_stud.middle_name[0].upper() + '. ' + \
                student[0].undergrad_stud.last_name.upper()
        else:
            full_name = address + ' ' + student[0].undergrad_stud.first_name.upper() + ' ' + \
                student[0].undergrad_stud.last_name.upper()

        address_stud = address + ' ' + student[0].undergrad_stud.last_name

        course = student[0].degree_courses.course_name + ' (' + \
            student[0].degree_courses.course_abbrev.upper() + ')'

        uni_registrar = db((db.registrar.registrar_position == 'University Registrar')).select(
            db.registrar.registrar_name)

        registrar = uni_registrar[0].registrar_name.upper()

        return locals()
    except IndexError:
        pass


def page3():
    # get student id from URL
    try:
        stud_arg = int(request.args(0))
    except TypeError:
        redirect(URL('index.html'))

    try:
        current_day = custom_strftime('{S}', dt.now())
        current_month_year = custom_strftime('%B, %Y', dt.now())

        student = db(db.undergrad_stud.student_id == stud_arg).select(
            db.undergrad_stud.first_name,
            db.undergrad_stud.middle_name,
            db.undergrad_stud.last_name,
            db.undergrad_stud.gender,
            db.undergrad_stud.date_graduated,
            db.undergrad_stud.gwa,
            db.degree_courses.course_name,
            db.degree_courses.course_abbrev,
            db.majors.major,
            left=[db.degree_courses.on(db.undergrad_stud.course_id ==
                                       db.degree_courses.course_id),
                  db.majors.on(db.undergrad_stud.major_id == db.majors.major_id)])

        if student[0].undergrad_stud.date_graduated is not None:
            date_graduated = custom_strftime('%B %d, %Y', student[
                0].undergrad_stud.date_graduated)
            year_graduated = custom_strftime('%Y', student[
                0].undergrad_stud.date_graduated)
        else:
            date_graduated = "N/A"
            year_graduated = "N/A"

        if student[0].undergrad_stud.middle_name is not None:
            full_name = student[0].undergrad_stud.first_name.upper() + ' ' + \
                student[0].undergrad_stud.middle_name[0].upper() + '. ' + \
                student[0].undergrad_stud.last_name.upper()
        else:
            full_name = student[0].undergrad_stud.first_name.upper() + ' ' + \
                student[0].undergrad_stud.last_name.upper()

        course = student[0].degree_courses.course_name + \
            ' (' + student[0].degree_courses.course_abbrev.upper() + ')'

        if student[0].majors.major is not None:
            major = student[0].majors.major
        else:
            major = "N/A"

        if student[0].undergrad_stud.gwa is not None:
            gwa = student[0].undergrad_stud.gwa
        else:
            gwa = "N/A"

        uni_registrar = db((db.registrar.registrar_position == 'University Registrar')).select(
            db.registrar.registrar_name)

        registrar = uni_registrar[0].registrar_name.upper()

        return locals()
    except IndexError:
        pass


def page4():
    # get student id from URL
    try:
        stud_arg = int(request.args(0))
    except TypeError:
        redirect(URL('index.html'))

    try:
        current_day = custom_strftime('{S}', dt.now())
        current_month_year = custom_strftime('%B, %Y', dt.now())

        student = db(db.undergrad_stud.student_id == stud_arg).select(
            db.undergrad_stud.first_name,
            db.undergrad_stud.middle_name,
            db.undergrad_stud.last_name,
            db.undergrad_stud.gender,
            db.undergrad_stud.date_graduated,
            db.undergrad_stud.award_id,
            db.degree_courses.course_name,
            db.degree_courses.course_abbrev,
            db.majors.major,
            db.awards.award_title,
            left=[db.degree_courses.on(db.undergrad_stud.course_id ==
                                       db.degree_courses.course_id),
                  db.majors.on(db.undergrad_stud.major_id ==
                               db.majors.major_id),
                  db.awards.on(db.undergrad_stud.award_id == db.awards.award_id)])

        if student[0].undergrad_stud.date_graduated is not None:
            date_graduated = custom_strftime(
                '%B %d, %Y', student[0].undergrad_stud.date_graduated)
            year_graduated = custom_strftime('%Y', student[
                0].undergrad_stud.date_graduated)
        else:
            date_graduated = "N/A"
            year_graduated = "N/A"

        if student[0].undergrad_stud.gender == 'F':
            address = 'Ms.'
            pronoun_1 = 'she'
            pronoun_2 = 'her'
        else:
            address = 'Mr.'
            pronoun_1 = 'he'
            pronoun_2 = 'his'

        if student[0].undergrad_stud.middle_name is not None:
            full_name = address + ' ' + student[0].undergrad_stud.first_name.upper() + \
                ' ' + student[0].undergrad_stud.middle_name[0].upper() + '. ' + \
                student[0].undergrad_stud.last_name.upper()
        else:
            full_name = address + ' ' + student[0].undergrad_stud.first_name.upper() + ' ' + \
                student[0].undergrad_stud.last_name.upper()

        address_stud = address + ' ' + student[0].undergrad_stud.last_name

        course = student[0].degree_courses.course_name + \
            ' (' + student[0].degree_courses.course_abbrev.upper() + ')'

        if student[0].majors.major is not None:
            major = student[0].majors.major
        else:
            major = "N/A"

        if student[0].undergrad_stud.award_id is not None:
            award = student[0].awards.award_title.upper()
        else:
            award = "N/A"

        uni_registrar = db((db.registrar.registrar_position == 'University Registrar')).select(
            db.registrar.registrar_name)

        registrar = uni_registrar[0].registrar_name.upper()

        return locals()
    except IndexError:
        pass


def page5():
    # get student id from URL
    try:
        stud_arg = int(request.args(0))
    except TypeError:
        redirect(URL('index.html'))

    try:
        current_day = custom_strftime('{S}', dt.now())
        current_month_year = custom_strftime('%B, %Y', dt.now())

        student = db((db.undergrad_stud.student_id == stud_arg)).select(
            db.undergrad_stud.first_name,
            db.undergrad_stud.middle_name,
            db.undergrad_stud.last_name,
            db.undergrad_stud.gender,
            db.undergrad_stud.date_graduated,
            db.degree_courses.course_name,
            db.degree_courses.course_abbrev,
            db.majors.major,
            db.campus_college_institute.cci_name,
            db.campus_college_institute.address,
            db.campus_college_institute.tel_no,
            left=[db.degree_courses.on(db.undergrad_stud.course_id ==
                                       db.degree_courses.course_id),
                  db.majors.on(db.undergrad_stud.major_id ==
                               db.majors.major_id),
                  db.campus_college_institute.on(db.degree_courses.cci_id ==
                                                 db.campus_college_institute.cci_id)])

        if student[0].undergrad_stud.date_graduated is not None:
            date_graduated = custom_strftime(
                '%B %d, %Y', student[0].undergrad_stud.date_graduated)
            year_graduated = custom_strftime('%Y', student[
                0].undergrad_stud.date_graduated)
        else:
            date_graduated = "N/A"
            year_graduated = "N/A"

        if student[0].undergrad_stud.gender == 'F':
            address = 'Ms.'
        else:
            address = 'Mr.'

        if student[0].undergrad_stud.middle_name is not None:
            full_name = address + ' ' + student[0].undergrad_stud.first_name.upper() + \
                ' ' + student[0].undergrad_stud.middle_name[0].upper() + '. ' + \
                student[0].undergrad_stud.last_name.upper()
        else:
            full_name = address + ' ' + student[0].undergrad_stud.first_name.upper() + ' ' + \
                student[0].undergrad_stud.last_name.upper()

        address_stud = address + ' ' + student[0].undergrad_stud.last_name

        course = student[0].degree_courses.course_name + \
            ' (' + student[0].degree_courses.course_abbrev.upper() + ')'

        if student[0].majors.major is not None:
            major = student[0].majors.major
        else:
            major = "N/A"

        registrar = db((db.registrar.registrar_position == 'Registrar II')).select(
            db.registrar.registrar_name)

        uni_registrar = db((db.registrar.registrar_position == 'University Registrar')).select(
            db.registrar.registrar_name)

        return locals()
    except IndexError:
        pass


def page6():
    # get student id from URL
    try:
        stud_arg = int(request.args(0))
    except TypeError:
        redirect(URL('index.html'))

    try:
        current_date = str(datetime.datetime.today().strftime('%B %d, %Y'))

        student = db((db.undergrad_stud.student_id == stud_arg)).select(
            db.undergrad_stud.first_name,
            db.undergrad_stud.middle_name,
            db.undergrad_stud.last_name,
            db.undergrad_stud.gender,
            db.undergrad_stud.date_graduated,
            db.degree_courses.course_name,
            db.degree_courses.course_abbrev,
            db.majors.major,
            db.campus_college_institute.cci_name,
            db.campus_college_institute.address,
            db.campus_college_institute.tel_no,
            left=[db.degree_courses.on(db.undergrad_stud.course_id ==
                                       db.degree_courses.course_id),
                  db.majors.on(db.undergrad_stud.major_id ==
                               db.majors.major_id),
                  db.campus_college_institute.on(db.degree_courses.cci_id ==
                                                 db.campus_college_institute.cci_id)])

        receipt = db((db.receipt.student_id == stud_arg)).select(
            db.receipt.or_no,
            db.receipt.date,
            db.receipt.amount,
            db.semester.sem,
            db.semester.acad_year,
            left=[db.semester.on(db.receipt.sem_id == db.semester.sem_id),
                  db.undergrad_stud.on(db.receipt.student_id ==
                                       db.undergrad_stud.student_id)])

        if student[0].undergrad_stud.gender == 'F':
            address = 'Ms.'
        else:
            address = 'Mr.'

        if student[0].undergrad_stud.middle_name is not None:
            full_name = address + ' ' + student[0].undergrad_stud.first_name.upper() + \
                ' ' + student[0].undergrad_stud.middle_name[0].upper() + '. ' + \
                student[0].undergrad_stud.last_name.upper()
        else:
            full_name = address + ' ' + student[0].undergrad_stud.first_name.upper() + ' ' + \
                student[0].undergrad_stud.last_name.upper()

        course = student[0].degree_courses.course_name + \
            ' (' + student[0].degree_courses.course_abbrev.upper() + ')'

        try:
            or_no = f'{receipt[0].receipt.or_no:07d}'
            receipt_date = custom_strftime('%m-%d-%y', receipt[0].receipt.date)
            last_term_enrolled = receipt[0].semester.sem[:7] + \
                '. SY ' + receipt[0].semester.acad_year
            receipt_amount = receipt[0].receipt.amount
        except IndexError:
            or_no = '<OR no.>'
            receipt_date = '<date>'
            last_term_enrolled = '<last term enrolled>'
            receipt_amount = '<amount>'

        registrar = db(
            (db.registrar.registrar_position == 'University Registrar')).select(
            db.registrar.registrar_position,
            db.registrar.registrar_name
        )

        return locals()
    except IndexError:
        pass
