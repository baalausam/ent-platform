import random
import json
from datetime import timedelta

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

from django import forms as dj_forms
from django.shortcuts import render, redirect, reverse
from django.contrib.auth.models import User
from django.contrib.auth import login, authenticate, logout as auth_logout
from django.contrib.auth.decorators import login_required
from django.utils import timezone
from django.utils.translation import gettext_lazy as _
from django.views.decorators.http import require_POST
from django.http import JsonResponse, HttpResponse

from .forms import (
    DirectorRegistrationForm, TeacherRegistrationForm,
    StudentRegistrationForm, QuestionForm, ExamSessionForm,
    QuestionGroupForm,
)
from .models import (
    School, SchoolClass, UserProfile, Subject,
    Question, TestAttempt, Answer, ExamSession,
    QuestionGroup,
    DIFFICULTY_A, DIFFICULTY_B, DIFFICULTY_C,
    TEST_DISTRIBUTION, CONTEXT_DIFFICULTY_COUNT,
)


# ============================================================
# ГЛАВНАЯ
# ============================================================

def home(request):
    if request.user.is_authenticated:
        try:
            profile = request.user.profile
            if profile.role == 'director':
                return redirect('core:director_dashboard')
            elif profile.role == 'zavuch':
                return redirect('core:zavuch_dashboard')
            elif profile.role == 'teacher':
                return redirect('core:teacher_dashboard')
            elif profile.role == 'student':
                return redirect('core:student_dashboard')
        except Exception:
            pass
    return render(request, 'core/home.html')


# ============================================================
# КӨМЕКШІ ФУНКЦИЯЛАР
# ============================================================

def _get_allowed_blocks(profile):
    codes = [s.code for s in profile.subjects.all()]
    allowed = []
    if 'kaz_history' in codes:
        allowed.append('kaz_history')
    if 'reading' in codes:
        allowed.append('reading')
    if 'math_literacy' in codes:
        allowed.append('math_literacy')
    if 'profile' in codes:
        allowed.append('profile')
    return allowed


def _can_create_groups(profile):
    return profile.subjects.filter(code='reading').exists() or \
           profile.subjects.filter(code='profile').exists()


def _get_active_session(profile):
    now = timezone.now()
    school = profile.school
    sessions = ExamSession.objects.filter(
        school=school, is_active=True,
        opens_at__lte=now, closes_at__gte=now,
    ).order_by('-created_at')
    for s in sessions:
        if s.classes.exists():
            if profile.school_class in s.classes.all():
                return s
        else:
            return s
    return None


# ============================================================
# ДИРЕКТОР
# ============================================================

def register_director(request):
    if request.method == 'POST':
        form = DirectorRegistrationForm(request.POST)
        if form.is_valid():
            data = form.cleaned_data
            user = User.objects.create_user(
                username=data['username'], password=data['password'],
                first_name=data['first_name'], last_name=data['last_name'],
            )
            school = School.objects.create(
                name=data['school_name'], city=data['city'], director=user,
            )
            UserProfile.objects.create(user=user, school=school, role='director')
            login(request, user)
            return redirect('core:director_setup_classes')
    else:
        form = DirectorRegistrationForm()
    return render(request, 'core/register_director.html', {'form': form})


@login_required
def director_setup_classes(request):
    profile = request.user.profile
    if profile.role != 'director':
        return redirect('core:home')
    school = profile.school
    if request.method == 'POST':
        selected = request.POST.getlist('classes')
        for item in selected:
            try:
                grade_str, letter = item.split('-')
                grade = int(grade_str)
                if grade in [9, 10, 11] and letter in ['А', 'Б', 'В', 'Г', 'Д', 'Е']:
                    SchoolClass.objects.get_or_create(
                        school=school, grade=grade, letter=letter,
                    )
            except (ValueError, IndexError):
                continue
        return redirect('core:director_dashboard')
    existing = set(school.classes.values_list('grade', 'letter'))
    return render(request, 'core/director_setup_classes.html', {
        'school': school, 'grades': [9, 10, 11],
        'letters': ['А', 'Б', 'В', 'Г', 'Д', 'Е'], 'existing': existing,
    })


@login_required
def director_dashboard(request):
    profile = request.user.profile
    if profile.role != 'director':
        return redirect('core:home')
    school = profile.school
    return render(request, 'core/director_dashboard.html', {
        'school': school, 'classes': school.classes.all(),
        'teacher_code': school.teacher_code,
        'zavuch_code': school.zavuch_code,
        'student_code': school.student_code,
    })


@login_required
def export_results_excel(request):
    profile = request.user.profile
    if profile.role not in ['director', 'zavuch']:
        return redirect('core:home')

    school = profile.school
    attempts = TestAttempt.objects.filter(
        student__profile__school=school
    ).select_related('student', 'student__profile', 'session').order_by('-started_at')

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Результаты"

    ws.append([
        '№', 'Ученик', 'Класс', 'Сессия',
        'Дата начала', 'Дата окончания',
        'Статус', 'Нарушений', 'Общий балл',
        'История КЗ', 'Чтение', 'Мат. грам.',
        'Профиль 1', 'Профиль 2',
    ])

    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="2C3E50", end_color="2C3E50", fill_type="solid")
        cell.alignment = Alignment(horizontal="center")

    for i, a in enumerate(attempts, 1):
        ws.append([
            i,
            f"{a.student.last_name} {a.student.first_name}",
            str(a.student.profile.school_class or '—'),
            a.session.title if a.session else '—',
            a.started_at.strftime('%d.%m.%Y %H:%M'),
            a.finished_at.strftime('%d.%m.%Y %H:%M') if a.finished_at else '—',
            a.get_status_display(),
            a.tab_switch_count,
            a.total_score,
            a.score_kaz_history,
            a.score_reading,
            a.score_math_literacy,
            a.score_profile_1,
            a.score_profile_2,
        ])

    widths = [5, 25, 8, 20, 18, 18, 15, 12, 12, 12, 10, 12, 12, 12]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    response = HttpResponse(
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    filename = f"results_{school.name}_{timezone.now().strftime('%Y%m%d')}.xlsx"
    response['Content-Disposition'] = f'attachment; filename="{filename}"'
    wb.save(response)
    return response


# ============================================================
# УЧИТЕЛЬ / ЗАВУЧ
# ============================================================

def register_teacher(request):
    if request.method == 'POST' and request.POST.get('step') == '2':
        first_name = request.POST.get('first_name', '').strip()
        last_name = request.POST.get('last_name', '').strip()
        username = request.POST.get('username', '').strip()
        password = request.POST.get('password', '')
        teacher_code = request.POST.get('teacher_code', '').strip().upper()
        role = request.POST.get('role', 'teacher')
        is_subject_teacher = request.POST.get('is_subject_teacher') == 'on'
        is_homeroom_teacher = request.POST.get('is_homeroom_teacher') == 'on'

        errors = []
        if not first_name: errors.append(_("Введите имя."))
        if not last_name: errors.append(_("Введите фамилию."))
        if not username: errors.append(_("Введите логин."))
        if not password: errors.append(_("Введите пароль."))
        if User.objects.filter(username=username).exists():
            errors.append(_("Такой логин уже занят."))

        if role == 'zavuch':
            is_subject_teacher = False
            is_homeroom_teacher = False
        else:
            if not is_subject_teacher and not is_homeroom_teacher:
                errors.append(_("Выберите хотя бы одно: предметник или классный руководитель."))

        school = None
        if role == 'zavuch':
            try:
                school = School.objects.get(zavuch_code=teacher_code)
            except School.DoesNotExist:
                errors.append(_("Неверный код завуча."))
        else:
            try:
                school = School.objects.get(teacher_code=teacher_code)
            except School.DoesNotExist:
                errors.append(_("Неверный код учителя."))

        if is_homeroom_teacher and school:
            homeroom_id = request.POST.get('homeroom_class')
            if homeroom_id:
                try:
                    target_class = SchoolClass.objects.get(id=homeroom_id, school=school)
                    if UserProfile.objects.filter(
                        homeroom_class=target_class,
                        is_homeroom_teacher=True
                    ).exists():
                        errors.append(
                            _("У класса %(cls)s уже есть классный руководитель.") % {'cls': target_class}
                        )
                except SchoolClass.DoesNotExist:
                    errors.append(_("Класс не найден."))

        if errors:
            return render(request, 'core/register_teacher.html', {
                'errors': errors, 'school': school,
                'all_subjects': Subject.objects.filter(is_active=True),
                'all_classes': school.classes.all() if school else [],
                'form_data': request.POST,
                'role': role,
            })

        user = User.objects.create_user(
            username=username, password=password,
            first_name=first_name, last_name=last_name,
        )
        profile = UserProfile.objects.create(
            user=user, school=school, role=role,
            is_subject_teacher=is_subject_teacher,
            is_homeroom_teacher=is_homeroom_teacher,
        )
        if is_subject_teacher:
            subject_ids = request.POST.getlist('subjects')
            if subject_ids:
                profile.subjects.set(Subject.objects.filter(id__in=subject_ids))
        if is_homeroom_teacher:
            homeroom_id = request.POST.get('homeroom_class')
            if homeroom_id:
                try:
                    profile.homeroom_class = SchoolClass.objects.get(id=homeroom_id, school=school)
                    profile.save()
                except SchoolClass.DoesNotExist:
                    pass

        login(request, user)
        if role == 'zavuch':
            return redirect('core:zavuch_dashboard')
        return redirect('core:teacher_dashboard')

    if request.method == 'POST' and request.POST.get('step') == '1':
        form = TeacherRegistrationForm(request.POST)
        if form.is_valid():
            data = form.cleaned_data
            code = data['teacher_code'].strip().upper()
            school = None
            role = None
            if School.objects.filter(teacher_code=code).exists():
                school = School.objects.get(teacher_code=code)
                role = 'teacher'
            elif School.objects.filter(zavuch_code=code).exists():
                school = School.objects.get(zavuch_code=code)
                role = 'zavuch'
            else:
                form.add_error('teacher_code', _("Неверный код."))
                return render(request, 'core/register_teacher.html', {'form': form, 'step': 1})
            return render(request, 'core/register_teacher.html', {
                'school': school, 'role': role,
                'all_subjects': Subject.objects.filter(is_active=True) if role == 'teacher' else [],
                'all_classes': school.classes.all() if role == 'teacher' else [],
                'form_data': data,
            })
        return render(request, 'core/register_teacher.html', {'form': form, 'step': 1})

    form = TeacherRegistrationForm()
    return render(request, 'core/register_teacher.html', {'form': form, 'step': 1})


@login_required
def teacher_dashboard(request):
    profile = request.user.profile
    if profile.role != 'teacher':
        return redirect('core:home')

    homeroom_students = []
    if profile.is_homeroom_teacher and profile.homeroom_class:
        homeroom_students = UserProfile.objects.filter(
            school_class=profile.homeroom_class,
            role='student'
        ).select_related('user').order_by('user__last_name', 'user__first_name')

    return render(request, 'core/teacher_dashboard.html', {
        'profile': profile,
        'homeroom_students': homeroom_students,
    })


@login_required
def zavuch_dashboard(request):
    profile = request.user.profile
    if profile.role != 'zavuch':
        return redirect('core:home')
    school = profile.school
    return render(request, 'core/zavuch_dashboard.html', {
        'school': school, 'classes': school.classes.all(),
        'total_students': UserProfile.objects.filter(school=school, role='student').count(),
        'total_teachers': UserProfile.objects.filter(school=school, role='teacher').count(),
    })


# ============================================================
# УЧЕНИК
# ============================================================

def register_student(request):
    if request.method == 'POST' and request.POST.get('step') == '2':
        first_name = request.POST.get('first_name', '').strip()
        last_name = request.POST.get('last_name', '').strip()
        username = request.POST.get('username', '').strip()
        password = request.POST.get('password', '')
        student_code = request.POST.get('student_code', '').strip().upper()

        errors = []
        if not first_name: errors.append(_("Введите имя."))
        if not last_name: errors.append(_("Введите фамилию."))
        if not username: errors.append(_("Введите логин."))
        if not password: errors.append(_("Введите пароль."))
        if User.objects.filter(username=username).exists():
            errors.append(_("Такой логин уже занят."))

        try:
            school = School.objects.get(student_code=student_code)
        except School.DoesNotExist:
            return render(request, 'core/register_student.html', {
                'errors': [_("Неверный код школы.")], 'step': 1,
            })

        class_id = request.POST.get('school_class')
        school_class = None
        if not class_id:
            errors.append(_("Выберите класс."))
        else:
            try:
                school_class = SchoolClass.objects.get(id=class_id, school=school)
            except SchoolClass.DoesNotExist:
                errors.append(_("Такого класса нет."))

        subj1_id = request.POST.get('profile_subject_1')
        subj2_id = request.POST.get('profile_subject_2')
        subj1 = subj2 = None
        if not subj1_id or not subj2_id:
            errors.append(_("Выберите два профильных предмета."))
        elif subj1_id == subj2_id:
            errors.append(_("Профильные предметы должны быть разными."))
        else:
            subj1 = Subject.objects.filter(id=subj1_id).first()
            subj2 = Subject.objects.filter(id=subj2_id).first()

        if errors:
            return render(request, 'core/register_student.html', {
                'errors': errors, 'school': school,
                'all_classes': school.classes.all(),
                'profile_subjects': Subject.objects.filter(category='profile', is_active=True),
                'form_data': request.POST,
            })

        user = User.objects.create_user(
            username=username, password=password,
            first_name=first_name, last_name=last_name,
        )
        UserProfile.objects.create(
            user=user, school=school, role='student',
            school_class=school_class,
            profile_subject_1=subj1, profile_subject_2=subj2,
        )
        login(request, user)
        return redirect('core:student_dashboard')

    if request.method == 'POST' and request.POST.get('step') == '1':
        form = StudentRegistrationForm(request.POST)
        if form.is_valid():
            data = form.cleaned_data
            try:
                school = School.objects.get(student_code=data['student_code'].strip().upper())
                return render(request, 'core/register_student.html', {
                    'school': school, 'all_classes': school.classes.all(),
                    'profile_subjects': Subject.objects.filter(category='profile', is_active=True),
                    'form_data': data,
                })
            except School.DoesNotExist:
                form.add_error('student_code', _("Неверный код."))
        return render(request, 'core/register_student.html', {'form': form, 'step': 1})

    form = StudentRegistrationForm()
    return render(request, 'core/register_student.html', {'form': form, 'step': 1})


@login_required
def student_dashboard(request):
    profile = request.user.profile
    if profile.role != 'student':
        return redirect('core:home')

    attempts = TestAttempt.objects.filter(student=request.user).order_by('-started_at')

    # ✅ БАРЛЫҚ ашық сессиялар
    now = timezone.now()
    active_sessions_qs = ExamSession.objects.filter(
        school=profile.school,
        is_active=True,
        opens_at__lte=now,
        closes_at__gte=now,
    ).order_by('-created_at')

    # Әр сессия үшін: оқушы тапсырды ма?
    sessions_data = []
    for s in active_sessions_qs:
        # Класс бойынша тексеру
        if s.classes.exists():
            if profile.school_class not in s.classes.all():
                continue  # бұл класс үшін емес

        # Бұл сессияда тапсырды ма?
        finished = TestAttempt.objects.filter(
            student=request.user, session=s, status='finished'
        ).exists()
        in_progress = TestAttempt.objects.filter(
            student=request.user, session=s, status='in_progress'
        ).first()

        sessions_data.append({
            'session': s,
            'finished': finished,
            'in_progress': in_progress,
        })

    # Жалпы in_progress (кез келген сессияда)
    in_progress = TestAttempt.objects.filter(
        student=request.user, status='in_progress'
    ).first()

    context = {
        'profile': profile,
        'school_class': profile.school_class,
        'subj1': profile.profile_subject_1,
        'subj2': profile.profile_subject_2,
        'attempts': attempts,
        'has_attempts': attempts.exists(),
        'active_sessions': sessions_data,  # ✅ ЖАҢА
        'in_progress': in_progress,
    }
    return render(request, 'core/student_dashboard.html', context)
# ============================================================
# СҰРАҚ ҚОСУ
# ============================================================

@login_required
def add_question(request):
    profile = request.user.profile
    if profile.role != 'teacher' or not profile.is_subject_teacher:
        return redirect('core:home')

    allowed = _get_allowed_blocks(profile)
    if not allowed:
        return render(request, 'core/access_denied.html', {
            'message': _("У вас нет предметов для добавления вопросов.")
        })

    group = None
    group_id = request.GET.get('group') or request.POST.get('group_id')
    if group_id:
        group = QuestionGroup.objects.filter(id=group_id, author=request.user).first()
        if group and group.current_question_count >= group.expected_question_count:
            return render(request, 'core/access_denied.html', {
                'message': _(
                    "Бұл контекстке %(exp)s сұрақ қосу керек. Артық қосуға болмайды."
                ) % {'exp': group.expected_question_count}
            })

    if request.method == 'POST':
        form = QuestionForm(request.POST, request.FILES, user=request.user)
        form.fields['block'].choices = [
            (code, label) for code, label in QuestionForm.BLOCK_CHOICES
            if code in allowed
        ]

        if form.is_valid():
            data = form.cleaned_data
            block = data['block']
            qtype = data['question_type']

            if block not in allowed:
                form.add_error('block', _("Вы не можете добавлять вопросы в этот блок."))
                return render(request, 'core/add_question.html', {
                    'form': form, 'profile': profile, 'group': group,
                })

            subject = data['subject'] if block == 'profile' else None

            if block == 'profile':
                if not subject or not profile.subjects.filter(id=subject.id).exists():
                    form.add_error('subject', _("Только по своему предмету."))
                    return render(request, 'core/add_question.html', {
                        'form': form, 'profile': profile, 'group': group,
                    })

            matching_data = None
            if qtype == 'matching':
                matching_data = _collect_matching_data(request.POST)
                if not matching_data:
                    form.add_error(None, _("Сәйкестендіру: 4 нұсқа + 2 кіші сұрақ қажет."))
                    return render(request, 'core/add_question.html', {
                        'form': form, 'profile': profile, 'group': group,
                    })

            Question.objects.create(
                block=block, subject=subject, group=group,
                question_type=qtype,
                kind='context' if group else 'standard',
                difficulty=data['difficulty'],
                text=data['text'],
                option_a=data.get('option_a', '') or '',
                option_b=data.get('option_b', '') or '',
                option_c=data.get('option_c', '') or '',
                option_d=data.get('option_d', '') or '',
                option_e=data.get('option_e', '') or '',
                option_f=data.get('option_f', '') or '',
                correct_answer=(data.get('correct_answer') or '').strip().upper(),
                matching_data=matching_data,
                image=data.get('image'),
                author=request.user, school=profile.school, status='draft',
            )

            if group:
                group.refresh_from_db()
                if group.current_question_count >= group.expected_question_count:
                    return redirect('core:my_question_groups')
                return redirect(f"{reverse('core:add_question')}?group={group.id}")

            return redirect('core:my_questions')
    else:
        form = QuestionForm(user=request.user)
        form.fields['block'].choices = [
            (code, label) for code, label in QuestionForm.BLOCK_CHOICES
            if code in allowed
        ]
        if len(allowed) == 1:
            form.fields['block'].initial = allowed[0]
            form.fields['block'].widget = dj_forms.HiddenInput()
        if group:
            form.fields['difficulty'].initial = group.difficulty

    return render(request, 'core/add_question.html', {
        'form': form, 'profile': profile, 'group': group,
    })


def _collect_matching_data(post_data):
    """Бір сәйкестендіру сұрағын жинау (4 нұсқа + 2 кіші сұрақ)."""
    options = [post_data.get(f'match_opt_{i}', '').strip() for i in range(4)]
    if not all(options):
        return None

    subs = []
    for sub_idx in range(2):
        sub_text = post_data.get(f'match_sub_{sub_idx}_text', '').strip()
        correct_idx = post_data.get(f'match_sub_{sub_idx}_correct', '')
        if not sub_text or correct_idx == '':
            return None
        try:
            correct = options[int(correct_idx)]
        except (ValueError, IndexError):
            return None
        subs.append({'text': sub_text, 'correct': correct})

    return {'options': options, 'sub_questions': subs}


@login_required
def my_questions(request):
    profile = request.user.profile
    if profile.role != 'teacher' or not profile.is_subject_teacher:
        return redirect('core:home')
    questions = Question.objects.filter(
        author=request.user, is_deleted_by_author=False
    ).order_by('-created_at')
    stats = {
        'total': questions.count(),
        'draft': questions.filter(status='draft').count(),
        'pending': questions.filter(status='pending').count(),
        'approved': questions.filter(status='approved').count(),
        'rejected': questions.filter(status='rejected').count(),
    }
    return render(request, 'core/my_questions.html', {
        'questions': questions, 'stats': stats, 'profile': profile,
    })


@login_required
def submit_question_for_review(request, question_id):
    profile = request.user.profile
    if profile.role != 'teacher':
        return redirect('core:home')
    try:
        q = Question.objects.get(id=question_id, author=request.user, status='draft',
                                 is_deleted_by_author=False)
        q.status = 'pending'
        q.save()
    except Question.DoesNotExist:
        pass
    return redirect('core:my_questions')


@login_required
@require_POST
def cancel_question(request, question_id):
    profile = request.user.profile
    if profile.role != 'teacher':
        return redirect('core:home')
    try:
        q = Question.objects.get(
            id=question_id, author=request.user,
            status__in=['draft', 'pending'],
            is_deleted_by_author=False,
        )
        q.is_deleted_by_author = True
        q.status = 'draft'
        q.save()

        if q.group:
            group = q.group
            if not group.is_complete:
                group.status = 'draft'
                group.save()
                group.questions.filter(
                    is_deleted_by_author=False
                ).update(status='draft')
    except Question.DoesNotExist:
        pass
    return redirect('core:my_questions')


@login_required
@require_POST
def cancel_group(request, group_id):
    profile = request.user.profile
    if profile.role != 'teacher':
        return redirect('core:home')
    try:
        g = QuestionGroup.objects.get(
            id=group_id, author=request.user,
            status__in=['draft', 'pending'],
        )
        g.questions.update(is_deleted_by_author=True)
        g.delete()
    except QuestionGroup.DoesNotExist:
        pass
    return redirect('core:my_question_groups')


# ============================================================
# КОНТЕКСТ
# ============================================================

@login_required
def add_question_group(request):
    profile = request.user.profile
    if profile.role != 'teacher' or not profile.is_subject_teacher:
        return redirect('core:home')

    allowed_blocks = _get_allowed_blocks(profile)
    can_create_reading = 'reading' in allowed_blocks
    can_create_profile = 'profile' in allowed_blocks

    if not can_create_reading and not can_create_profile:
        return render(request, 'core/access_denied.html', {
            'message': _("Контексті тек оқу сауаттылығы немесе бейіндік пән мұғалімі қоса алады.")
        })

    if request.method == 'POST':
        form = QuestionGroupForm(request.POST)
        if form.is_valid():
            group = form.save(commit=False)
            group.school = profile.school
            group.author = request.user

            block_choice = request.POST.get('block_choice', '')

            if block_choice == 'profile' and can_create_profile:
                group.block = 'profile'
                group.subject = profile.subjects.filter(category='profile').first()
                group.difficulty = 'A'
            elif block_choice == 'reading' and can_create_reading:
                group.block = 'reading'
                group.subject = None
            else:
                if can_create_profile and not can_create_reading:
                    group.block = 'profile'
                    group.subject = profile.subjects.filter(category='profile').first()
                    group.difficulty = 'A'
                else:
                    group.block = 'reading'
                    group.subject = None

            group.save()
            return redirect(f"{reverse('core:add_question')}?group={group.id}")
    else:
        form = QuestionGroupForm()

    return render(request, 'core/add_question_group.html', {
        'form': form, 'profile': profile,
        'can_create_reading': can_create_reading,
        'can_create_profile': can_create_profile,
    })


@login_required
def my_question_groups(request):
    profile = request.user.profile
    if profile.role != 'teacher':
        return redirect('core:home')
    if not _can_create_groups(profile):
        return render(request, 'core/access_denied.html', {
            'message': _("Контексты доступны только учителю грамотности чтения или профильного предмета.")
        })
    groups = QuestionGroup.objects.filter(author=request.user).order_by('-created_at')
    return render(request, 'core/my_question_groups.html', {
        'groups': groups, 'profile': profile,
    })


@login_required
@require_POST
def submit_group_for_review(request, group_id):
    profile = request.user.profile
    if profile.role != 'teacher':
        return redirect('core:home')
    try:
        g = QuestionGroup.objects.get(id=group_id, author=request.user, status='draft')

        if not g.is_complete:
            return render(request, 'core/access_denied.html', {
                'message': _(
                    "Контекст толық емес: %(cur)s / %(exp)s сұрақ. "
                    "Алдымен барлық сұрақтарды қосыңыз."
                ) % {'cur': g.current_question_count, 'exp': g.expected_question_count}
            })

        non_draft = g.questions.filter(
            is_deleted_by_author=False
        ).exclude(status='draft').count()
        if non_draft > 0:
            return render(request, 'core/access_denied.html', {
                'message': _("Контексте бекітілмеген сұрақтар бар.")
            })

        g.status = 'pending'
        g.save()
        g.questions.filter(
            status='draft', is_deleted_by_author=False
        ).update(status='pending')
    except QuestionGroup.DoesNotExist:
        pass
    return redirect('core:my_question_groups')


# ============================================================
# ТЕКСЕРУ
# ============================================================

def _is_admin_role(profile):
    return profile.role in ['director', 'zavuch']


@login_required
def review_questions(request):
    profile = request.user.profile
    if not _is_admin_role(profile):
        return redirect('core:home')
    school = profile.school
    status_filter = request.GET.get('status', 'pending')

    all_qs = Question.objects.filter(
        school=school, is_deleted_by_author=False
    ).order_by('-created_at')

    groups_pending_all = QuestionGroup.objects.filter(
        school=school, status='pending'
    ).prefetch_related('questions').order_by('-created_at')

    groups_pending = [g for g in groups_pending_all if g.is_complete]

    if status_filter == 'all':
        questions = all_qs
    elif status_filter in ['draft', 'pending', 'approved', 'rejected']:
        questions = all_qs.filter(status=status_filter)
    else:
        questions = all_qs.filter(status='pending')

    stats = {
        'draft': all_qs.filter(status='draft').count(),
        'pending': all_qs.filter(status='pending').count(),
        'approved': all_qs.filter(status='approved').count(),
        'rejected': all_qs.filter(status='rejected').count(),
        'all': all_qs.count(),
    }
    return render(request, 'core/review_questions.html', {
        'questions': questions, 'stats': stats,
        'status_filter': status_filter, 'profile': profile,
        'groups_pending': groups_pending,
    })


@login_required
def approve_question(request, question_id):
    profile = request.user.profile
    if not _is_admin_role(profile):
        return redirect('core:home')
    try:
        q = Question.objects.get(id=question_id, school=profile.school, status='pending')
        q.status = 'approved'
        q.approved_by = request.user
        q.approved_at = timezone.now()
        q.save()
    except Question.DoesNotExist:
        pass
    return redirect('core:review_questions')


@login_required
def reject_question(request, question_id):
    profile = request.user.profile
    if not _is_admin_role(profile):
        return redirect('core:home')
    if request.method == 'POST':
        try:
            q = Question.objects.get(id=question_id, school=profile.school, status='pending')
            q.status = 'rejected'
            q.approved_by = request.user
            q.save()
        except Question.DoesNotExist:
            pass
    return redirect('core:review_questions')


@login_required
@require_POST
def approve_group(request, group_id):
    profile = request.user.profile
    if not _is_admin_role(profile):
        return redirect('core:home')
    try:
        g = QuestionGroup.objects.get(id=group_id, school=profile.school, status='pending')
        g.status = 'approved'
        g.approved_by = request.user
        g.approved_at = timezone.now()
        g.save()
        g.questions.filter(status='pending').update(status='approved')
    except QuestionGroup.DoesNotExist:
        pass
    return redirect('core:review_questions')


@login_required
@require_POST
def reject_group(request, group_id):
    profile = request.user.profile
    if not _is_admin_role(profile):
        return redirect('core:home')
    try:
        g = QuestionGroup.objects.get(id=group_id, school=profile.school, status='pending')
        g.status = 'rejected'
        g.save()
        g.questions.filter(status='pending').update(status='rejected')
    except QuestionGroup.DoesNotExist:
        pass
    return redirect('core:review_questions')


# ============================================================
# ТЕСТ — ІРІКТЕУ
# ============================================================

def _pick_by_difficulty(school, block, subject, difficulty, count,
                        qtype=None, kind=None):
    qs = Question.objects.filter(
        school=school, block=block, status='approved',
        difficulty=difficulty, is_deleted_by_author=False,
    )
    if subject:
        qs = qs.filter(subject=subject)
    if qtype:
        qs = qs.filter(question_type=qtype)
    if kind:
        qs = qs.filter(kind=kind)
    qs = list(qs)
    if len(qs) < count:
        return qs
    return random.sample(qs, count)


def _pick_reading_group(school, difficulty):
    """Оқу сауаттылығы үшін деңгейге сай толық контекст таңдау."""
    groups = QuestionGroup.objects.filter(
        school=school, block='reading', status='approved', difficulty=difficulty,
    )
    valid = [g for g in groups if g.is_complete]
    if not valid:
        return None
    return random.choice(valid)


def _build_profile_questions(school, subject):
    """
    Бейіндік пән: 40 сұрақ = 20A + 12B + 8C (ЖАЛПЫ)
    Тәртіп: single → context → matching → multiple
    """
    result = []

    # ============ 1-25: SINGLE (25 сұрақ) ============
    singles = list(Question.objects.filter(
        school=school, block='profile', subject=subject,
        status='approved', question_type='single', kind='standard',
        is_deleted_by_author=False,
    ))
    if len(singles) < 25:
        result += singles
    else:
        result += random.sample(singles, 25)

    # ============ 26-30: CONTEXT (5 сұрақ) ============
    ctx_groups = QuestionGroup.objects.filter(
        school=school, block='profile', subject=subject, status='approved',
    )
    valid_ctx = [g for g in ctx_groups if g.current_question_count >= 5]
    if valid_ctx:
        chosen_ctx = random.choice(valid_ctx)
        result += list(chosen_ctx.questions.filter(
            status='approved', is_deleted_by_author=False
        ).order_by('id')[:5])

    # ============ 31-35: MATCHING (5 сұрақ) ============
    matchings = list(Question.objects.filter(
        school=school, block='profile', subject=subject,
        status='approved', question_type='matching',
        is_deleted_by_author=False,
    ))
    if len(matchings) < 5:
        result += matchings
    else:
        result += random.sample(matchings, 5)

    # ============ 36-40: MULTIPLE (5 сұрақ) ============
    multiples = list(Question.objects.filter(
        school=school, block='profile', subject=subject,
        status='approved', question_type='multiple',
        is_deleted_by_author=False,
    ))
    if len(multiples) < 5:
        result += multiples
    else:
        result += random.sample(multiples, 5)

    return result  # 40 сұрақ


def _analyze_profile_questions(questions):
    """
    Профиль сұрақтарының құрылымын талдау.
    25 single + 5 ctx + 5 matching + 5 multiple = 40.
    """
    details = {
        'single': 0,
        'context': 0,
        'matching': 0,
        'multiple': 0,
    }
    for q in questions:
        if q.kind == 'context':
            details['context'] += 1
        elif q.question_type == 'single':
            details['single'] += 1
        elif q.question_type == 'matching':
            details['matching'] += 1
        elif q.question_type == 'multiple':
            details['multiple'] += 1
    return details


def _subject_label(q):
    if q.block == 'kaz_history':
        return 'kaz_history'
    if q.block == 'reading':
        return 'reading'
    if q.block == 'math_literacy':
        return 'math_literacy'
    if q.block == 'profile':
        name = q.subject.name_ru if q.subject else 'unknown'
        return f"profile_{name}"
    return q.block


# ============================================================
# ТЕСТ БАСТАУ
# ============================================================

@login_required
def start_test(request):
    profile = request.user.profile
    if profile.role != 'student':
        return redirect('core:home')

    # 1. Белсенді (in_progress) тест бар ма?
    active = TestAttempt.objects.filter(student=request.user, status='in_progress').first()
    if active:
        return redirect('core:take_test', attempt_id=active.id)

    # 2. ✅ Сессияны анықтау
    session_id = request.GET.get('session')
    now = timezone.now()

    if session_id:
        # Оқушы нақты сессия таңдады
        try:
            available = ExamSession.objects.get(
                id=session_id,
                school=profile.school,
                is_active=True,
            )
        except ExamSession.DoesNotExist:
            return redirect('core:student_dashboard')

        # Уақыт тексеру
        if not (available.opens_at <= now <= available.closes_at):
            return render(request, 'core/no_active_session.html', {'next_session': None})

        # Класс бойынша тексеру
        if available.classes.exists():
            if profile.school_class not in available.classes.all():
                return redirect('core:student_dashboard')
    else:
        # Автоматты — ең соңғы ашық сессия
        available = _get_active_session(profile)
        if not available:
            next_session = ExamSession.objects.filter(
                school=profile.school, is_active=True, opens_at__gt=now
            ).order_by('opens_at').first()
            return render(request, 'core/no_active_session.html', {'next_session': next_session})

    # 3. Бұл сессияда бұрын тапсырған ба?
    if TestAttempt.objects.filter(
        student=request.user, session=available, status='finished'
    ).exists():
        return render(request, 'core/test_already_done.html')

    # 4. Профильдік пәндер
    subj1 = profile.profile_subject_1
    subj2 = profile.profile_subject_2
    if not subj1 or not subj2:
        return render(request, 'core/test_no_subjects.html')

    # ============================================================
    # 5 БЛОКТЫ ЖЕКЕ ЖИНАУ
    # ============================================================

    # 1. ҚАЗАҚСТАН ТАРИХЫ — 20
    q_kaz = []
    kaz_details = {'A': 0, 'B': 0, 'C': 0}
    for diff, need in [('A', 10), ('B', 6), ('C', 4)]:
        picked = _pick_by_difficulty(profile.school, 'kaz_history', None, diff, need)
        kaz_details[diff] = len(picked)
        q_kaz += picked

    # 2. ОҚУ САУАТТЫЛЫҒЫ — 10 (3 контекст)
    q_read = []
    reading_groups_used = []
    read_details = {'A': 0, 'B': 0, 'C': 0}
    for diff in ['A', 'B', 'C']:
        g = _pick_reading_group(profile.school, diff)
        if g:
            questions = list(g.questions.filter(
                status='approved', is_deleted_by_author=False
            ).order_by('id'))
            read_details[diff] = len(questions)
            q_read += questions
            reading_groups_used.append(g)

    # 3. МАТЕМАТИКАЛЫҚ САУАТТЫЛЫҚ — 10
    q_math = []
    math_details = {'A': 0, 'B': 0, 'C': 0}
    for diff, need in [('A', 5), ('B', 3), ('C', 2)]:
        picked = _pick_by_difficulty(profile.school, 'math_literacy', None, diff, need)
        math_details[diff] = len(picked)
        q_math += picked

    # 4. ПРОФИЛЬ 1 — 40
    q_p1 = _build_profile_questions(profile.school, subj1)
    p1_details = _analyze_profile_questions(q_p1)

    # 5. ПРОФИЛЬ 2 — 40
    q_p2 = _build_profile_questions(profile.school, subj2)
    p2_details = _analyze_profile_questions(q_p2)

    all_questions = q_kaz + q_read + q_math + q_p1 + q_p2

    # ============================================================
    # ҚАТАҢ ТЕКСЕРУ
    # ============================================================
    errors = []

    if len(q_kaz) < 20:
        errors.append({
            'block': 'Қазақстан тарихы',
            'have': len(q_kaz), 'need': 20,
            'details': kaz_details,
        })
    if len(q_read) < 10:
        errors.append({
            'block': 'Оқу сауаттылығы',
            'have': len(q_read), 'need': 10,
            'details': read_details,
            'groups': len(reading_groups_used),
        })
    if len(q_math) < 10:
        errors.append({
            'block': 'Математикалық сауаттылық',
            'have': len(q_math), 'need': 10,
            'details': math_details,
        })
    if len(q_p1) < 40:
        errors.append({
            'block': f'Профиль 1 — {subj1.name_ru}',
            'have': len(q_p1), 'need': 40,
            'details': p1_details,
        })
    if len(q_p2) < 40:
        errors.append({
            'block': f'Профиль 2 — {subj2.name_ru}',
            'have': len(q_p2), 'need': 40,
            'details': p2_details,
        })

    if errors:
        return render(request, 'core/test_not_enough.html', {
            'have': len(all_questions),
            'need': 120,
            'errors': errors,
            'subj1_name': subj1.name_ru,
            'subj2_name': subj2.name_ru,
        })

    # ============================================================
    # БАРЛЫҒЫ ДАЙЫН — тестті бастау
    # ============================================================
    attempt = TestAttempt.objects.create(
        student=request.user, status='in_progress',
        session=available,
        profile_subject_1=subj1.name_ru if subj1 else '',
        profile_subject_2=subj2.name_ru if subj2 else '',
    )

    for q in all_questions:
        Answer.objects.create(
            attempt=attempt, question=q,
            question_snapshot=q.text,
            correct_answer_snapshot=q.correct_answer,
            question_type_snapshot=q.question_type,
            subject_snapshot=_subject_label(q),
            difficulty_snapshot=q.difficulty,
            kind_snapshot=q.kind,
            matching_data_snapshot=q.matching_data,
            image_snapshot=q.image.url if q.image else '',
            group_context_snapshot=q.group.context_text if q.group else '',
            group_title_snapshot=q.group.title if q.group else '',
        )

    return redirect('core:take_test', attempt_id=attempt.id)

# ============================================================
# ТЕСТ ТАПСЫРУ
# ============================================================

@login_required
def take_test(request, attempt_id):
    try:
        attempt = TestAttempt.objects.get(id=attempt_id, student=request.user)
    except TestAttempt.DoesNotExist:
        return redirect('core:home')
    if attempt.status != 'in_progress':
        return redirect('core:test_result', attempt_id=attempt.id)
    elapsed = timezone.now() - attempt.started_at
    if elapsed > timedelta(hours=4):
        _finish_attempt(attempt)
        return redirect('core:test_result', attempt_id=attempt.id)

    answers = attempt.answers.select_related('question').order_by('id')

    answers_data = []
    for ans in answers:
        matching_rows = []
        if ans.question_type_snapshot == 'matching' and ans.matching_data_snapshot:
            subs = ans.matching_data_snapshot.get('sub_questions', [])
            student_ans = ans.matching_answers or {}
            options = ans.matching_data_snapshot.get('options', [])

            for sub_idx, sub in enumerate(subs):
                key = str(sub_idx)
                selected = student_ans.get(key, '')
                opts_with_selected = [
                    {'value': opt, 'selected': (opt == selected)}
                    for opt in options
                ]
                matching_rows.append({
                    'text': sub.get('text', ''),
                    'options': opts_with_selected,
                })

        answers_data.append({
            'ans': ans,
            'matching_rows': matching_rows,
        })

    return render(request, 'core/take_test.html', {
        'attempt': attempt,
        'answers': answers,
        'answers_data': answers_data,
        'seconds_left': int((timedelta(hours=4) - elapsed).total_seconds()),
    })


@login_required
@require_POST
def save_answer(request, attempt_id, answer_id):
    try:
        attempt = TestAttempt.objects.get(id=attempt_id, student=request.user, status='in_progress')
    except TestAttempt.DoesNotExist:
        return JsonResponse({'error': 'not_found'}, status=404)
    elapsed = timezone.now() - attempt.started_at
    if elapsed > timedelta(hours=4):
        _finish_attempt(attempt)
        return JsonResponse({'error': 'time_up'})

    try:
        answer = Answer.objects.get(id=answer_id, attempt=attempt)
    except Answer.DoesNotExist:
        return JsonResponse({'error': 'answer_not_found'}, status=404)

    if answer.question_type_snapshot == 'matching':
        try:
            matching_answers = json.loads(request.POST.get('matching', '{}'))
        except (json.JSONDecodeError, TypeError):
            matching_answers = {}
        answer.matching_answers = matching_answers
        answer.chosen = ''
        answer.save()
    else:
        chosen = request.POST.get('chosen', '').strip().upper()
        chosen = ''.join(sorted(set(c for c in chosen if c in 'ABCDEF')))
        answer.chosen = chosen
        answer.save()

    return JsonResponse({
        'ok': True,
        'answered': attempt.answers.exclude(chosen='', matching_answers=None).count(),
        'total': attempt.answers.count(),
    })


@login_required
@require_POST
def register_violation(request, attempt_id):
    try:
        attempt = TestAttempt.objects.get(id=attempt_id, student=request.user, status='in_progress')
    except TestAttempt.DoesNotExist:
        return JsonResponse({'error': 'not_found'}, status=404)
    attempt.tab_switch_count += 1
    MAX_VIOLATIONS = 3
    if attempt.tab_switch_count >= MAX_VIOLATIONS:
        attempt.status = 'annulled'
        attempt.finished_at = timezone.now()
        attempt.save()
        return JsonResponse({'count': attempt.tab_switch_count, 'max': MAX_VIOLATIONS, 'annulled': True})
    attempt.save()
    return JsonResponse({'count': attempt.tab_switch_count, 'max': MAX_VIOLATIONS, 'annulled': False})


@login_required
@require_POST
def finish_test(request, attempt_id):
    try:
        attempt = TestAttempt.objects.get(id=attempt_id, student=request.user, status='in_progress')
    except TestAttempt.DoesNotExist:
        return redirect('core:home')
    _finish_attempt(attempt)
    return redirect('core:test_result', attempt_id=attempt.id)


# ============================================================
# БАҒАЛАУ
# ============================================================

def _score_single(answer):
    correct = (answer.correct_answer_snapshot or '').upper()
    chosen = (answer.chosen or '').upper()
    if chosen and correct and chosen == correct:
        return 1
    return 0


def _score_multiple(answer):
    correct = (answer.correct_answer_snapshot or '').upper()
    chosen = (answer.chosen or '').upper()
    if not chosen or not correct:
        return 0
    chosen_set = set(chosen)
    correct_set = set(correct)
    if chosen_set == correct_set:
        return 2
    wrong = len(chosen_set - correct_set)
    right = len(chosen_set & correct_set)
    if wrong == 0 and right > 0:
        return 1
    if wrong == 1 and right >= 2:
        return 1
    return 0


def _score_matching(answer):
    md = answer.matching_data_snapshot or {}
    subs = md.get('sub_questions', [])
    student_ans = answer.matching_answers or {}

    correct_count = 0
    for sub_idx, sub in enumerate(subs):
        key = str(sub_idx)
        student_choice = student_ans.get(key, '')
        if student_choice and student_choice == sub.get('correct'):
            correct_count += 1

    if correct_count == 2:
        return 2
    elif correct_count == 1:
        return 1
    return 0


def _finish_attempt(attempt):
    kaz = read = math = 0

    for ans in attempt.answers.all():
        qtype = ans.question_type_snapshot
        subj = ans.subject_snapshot

        if qtype == 'single':
            score = _score_single(ans)
        elif qtype == 'multiple':
            score = _score_multiple(ans)
        elif qtype == 'matching':
            score = _score_matching(ans)
        else:
            score = 0

        ans.score = score
        ans.save()

        if subj == 'kaz_history':
            kaz += score
        elif subj == 'reading':
            read += score
        elif subj == 'math_literacy':
            math += score

    profile_answers = list(
        attempt.answers.filter(subject_snapshot__startswith='profile_').order_by('id')
    )
    p1 = sum(a.score for a in profile_answers[:40])
    p2 = sum(a.score for a in profile_answers[40:80])

    attempt.score_kaz_history = kaz
    attempt.score_reading = read
    attempt.score_math_literacy = math
    attempt.score_profile_1 = p1
    attempt.score_profile_2 = p2
    attempt.total_score = kaz + read + math + p1 + p2
    attempt.passed_thresholds = (kaz >= 5 and read >= 3 and math >= 3 and p1 >= 5 and p2 >= 5)
    attempt.status = 'finished'
    attempt.finished_at = timezone.now()
    attempt.save()


# ============================================================
# НӘТИЖЕ
# ============================================================

@login_required
def test_result(request, attempt_id):
    profile = request.user.profile
    try:
        attempt = TestAttempt.objects.select_related(
            'student', 'student__profile', 'session'
        ).prefetch_related(
            'answers__question'
        ).get(id=attempt_id)
    except TestAttempt.DoesNotExist:
        return redirect('core:home')

    view_mode = None
    if profile.role == 'student':
        if attempt.student == request.user:
            view_mode = 'student'
    elif profile.role in ['director', 'zavuch']:
        if attempt.student.profile.school == profile.school:
            view_mode = 'admin'
    elif profile.role == 'teacher':
        if profile.is_homeroom_teacher and attempt.student.profile.school_class == profile.homeroom_class:
            view_mode = 'homeroom'
    if not view_mode:
        return redirect('core:home')

    answers_with_matching = []
    for ans in attempt.answers.all().order_by('id'):
        matching_detail = []
        if ans.question_type_snapshot == 'matching' and ans.matching_data_snapshot:
            student_ans = ans.matching_answers or {}
            for sub_idx, sub in enumerate(ans.matching_data_snapshot.get('sub_questions', [])):
                key = str(sub_idx)
                matching_detail.append({
                    'text': sub.get('text', ''),
                    'chosen': student_ans.get(key, '—'),
                    'correct': sub.get('correct', ''),
                })
        answers_with_matching.append({
            'ans': ans,
            'matching_detail': matching_detail,
        })

    return render(request, 'core/test_result.html', {
        'attempt': attempt,
        'view_mode': view_mode,
        'student': attempt.student,
        'answers_with_matching': answers_with_matching,
    })


# ============================================================
# МЕКТЕП / СЫНЫП НӘТИЖЕЛЕРІ
# ============================================================

@login_required
def school_results(request):
    profile = request.user.profile
    if profile.role not in ['director', 'zavuch']:
        return redirect('core:home')

    school = profile.school
    sessions = ExamSession.objects.filter(
        school=school, attempts__isnull=False,
    ).distinct().order_by('-created_at')

    session_data = []
    for s in sessions:
        finished = TestAttempt.objects.filter(
            session=s, student__profile__school=school, status='finished'
        ).values('student').distinct().count()
        session_data.append({'session': s, 'finished': finished})

    return render(request, 'core/school_results.html', {
        'school': school, 'session_data': session_data, 'profile': profile,
    })


@login_required
def school_results_detail(request, session_id):
    profile = request.user.profile
    if profile.role not in ['director', 'zavuch']:
        return redirect('core:home')

    school = profile.school
    try:
        session = ExamSession.objects.get(id=session_id, school=school)
    except ExamSession.DoesNotExist:
        return redirect('core:school_results')

    class_id = request.GET.get('class')
    status_filter = request.GET.get('status', 'all')

    attempts = TestAttempt.objects.filter(
        session=session, student__profile__school=school,
    ).select_related('student', 'student__profile').order_by('-started_at')

    if class_id:
        attempts = attempts.filter(student__profile__school_class_id=class_id)
    if status_filter in ['finished', 'annulled', 'in_progress']:
        attempts = attempts.filter(status=status_filter)

    return render(request, 'core/school_results_detail.html', {
        'school': school, 'session': session, 'attempts': attempts,
        'classes': school.classes.all(), 'selected_class': class_id,
        'status_filter': status_filter,
        'total_count': attempts.values('student').distinct().count(),
    })


@login_required
def class_results(request):
    profile = request.user.profile
    if profile.role != 'teacher' or not profile.is_homeroom_teacher:
        return redirect('core:home')
    if not profile.homeroom_class:
        return render(request, 'core/class_results.html', {
            'no_class': True, 'profile': profile,
        })

    school_class = profile.homeroom_class
    sessions = ExamSession.objects.filter(
        attempts__student__profile__school_class=school_class
    ).distinct().order_by('-created_at')

    session_data = []
    for s in sessions:
        finished = TestAttempt.objects.filter(
            session=s, student__profile__school_class=school_class, status='finished'
        ).values('student').distinct().count()
        session_data.append({'session': s, 'finished': finished})

    return render(request, 'core/class_results.html', {
        'school_class': school_class, 'session_data': session_data, 'profile': profile,
    })


@login_required
def class_results_detail(request, session_id):
    profile = request.user.profile
    if profile.role != 'teacher' or not profile.is_homeroom_teacher:
        return redirect('core:home')
    if not profile.homeroom_class:
        return redirect('core:class_results')

    school_class = profile.homeroom_class
    try:
        session = ExamSession.objects.get(id=session_id)
    except ExamSession.DoesNotExist:
        return redirect('core:class_results')

    students = UserProfile.objects.filter(
        school_class=school_class, role='student'
    ).select_related('user').order_by('user__last_name', 'user__first_name')

    attempts = TestAttempt.objects.filter(
        student__profile__school_class=school_class, session=session,
    ).select_related('student').order_by('-started_at')

    attempts_by_student = {}
    for a in attempts:
        if a.student_id not in attempts_by_student:
            attempts_by_student[a.student_id] = a

    rows = [
        {'student': s.user, 'profile': s, 'attempt': attempts_by_student.get(s.user_id)}
        for s in students
    ]

    students_finished = attempts.filter(status='finished').values('student').distinct().count()

    return render(request, 'core/class_results_detail.html', {
        'school_class': school_class, 'session': session, 'rows': rows,
        'profile': profile, 'total_students': students.count(),
        'students_finished': students_finished,
    })


# ============================================================
# СЕССИЯ
# ============================================================

def _can_manage_sessions(profile):
    if profile.role in ['director', 'zavuch']:
        return True
    if profile.role == 'teacher' and profile.is_homeroom_teacher:
        return True
    return False


@login_required
def manage_sessions(request):
    profile = request.user.profile
    if not _can_manage_sessions(profile):
        return redirect('core:home')
    school = profile.school
    if request.method == 'POST':
        form = ExamSessionForm(request.POST, school=school)
        if form.is_valid():
            data = form.cleaned_data
            session = ExamSession.objects.create(
                school=school, title=data['title'],
                opens_at=data['opens_at'], closes_at=data['closes_at'],
                created_by=request.user,
            )
            session.classes.set(data['classes'])
            return redirect('core:manage_sessions')
    else:
        form = ExamSessionForm(school=school)
    sessions = ExamSession.objects.filter(school=school).order_by('-created_at')
    return render(request, 'core/manage_sessions.html', {
        'form': form, 'sessions': sessions, 'profile': profile,
    })


@login_required
@require_POST
def toggle_session(request, session_id):
    profile = request.user.profile
    if not _can_manage_sessions(profile):
        return redirect('core:home')
    try:
        s = ExamSession.objects.get(id=session_id, school=profile.school)
        s.is_active = not s.is_active
        s.save()
    except ExamSession.DoesNotExist:
        pass
    return redirect('core:manage_sessions')


@login_required
@require_POST
def delete_session(request, session_id):
    profile = request.user.profile
    if not _can_manage_sessions(profile):
        return redirect('core:home')
    ExamSession.objects.filter(id=session_id, school=profile.school).delete()
    return redirect('core:manage_sessions')


# ============================================================
# КІРУ / ШЫҒУ
# ============================================================

def login_view(request):
    if request.user.is_authenticated:
        try:
            return _redirect_by_role(request.user.profile.role)
        except Exception:
            pass
    error = None
    if request.method == 'POST':
        username = request.POST.get('username', '').strip()
        password = request.POST.get('password', '')
        user = authenticate(request, username=username, password=password)
        if user is not None:
            login(request, user)
            try:
                return _redirect_by_role(user.profile.role)
            except Exception:
                return redirect('core:home')
        else:
            error = _("Неверный логин или пароль.")
    return render(request, 'core/login.html', {'error': error})


def logout_view(request):
    auth_logout(request)
    return redirect('core:home')


def _redirect_by_role(role):
    if role == 'student':
        return redirect('core:student_dashboard')
    if role == 'teacher':
        return redirect('core:teacher_dashboard')
    if role == 'zavuch':
        return redirect('core:zavuch_dashboard')
    if role == 'director':
        return redirect('core:director_dashboard')
    return redirect('core:home')


# ============================================================
# СПИСКТЕР
# ============================================================

@login_required
def students_list(request):
    profile = request.user.profile
    if profile.role not in ['director', 'zavuch']:
        return redirect('core:home')
    school = profile.school
    class_id = request.GET.get('class')
    students = UserProfile.objects.filter(
        school=school, role='student'
    ).select_related('user', 'school_class').order_by(
        'school_class__grade', 'school_class__letter', 'user__last_name'
    )
    if class_id:
        students = students.filter(school_class_id=class_id)
    return render(request, 'core/students_list.html', {
        'school': school, 'students': students, 'classes': school.classes.all(),
        'selected_class': class_id, 'total': students.count(),
    })


@login_required
def teachers_list(request):
    profile = request.user.profile
    if profile.role not in ['director', 'zavuch']:
        return redirect('core:home')
    school = profile.school
    teachers = UserProfile.objects.filter(
        school=school, role='teacher'
    ).select_related('user', 'homeroom_class').prefetch_related('subjects').order_by('user__last_name')
    return render(request, 'core/teachers_list.html', {
        'school': school, 'teachers': teachers, 'total': teachers.count(),
    })


@login_required
def class_students_list(request, class_id):
    profile = request.user.profile
    if profile.role not in ['director', 'zavuch']:
        return redirect('core:home')

    school = profile.school
    try:
        school_class = SchoolClass.objects.get(id=class_id, school=school)
    except SchoolClass.DoesNotExist:
        return redirect('core:director_dashboard')

    students = UserProfile.objects.filter(
        school_class=school_class, role='student'
    ).select_related('user').order_by('user__last_name', 'user__first_name')

    return render(request, 'core/class_students_list.html', {
        'school': school, 'school_class': school_class, 'students': students,
        'profile': profile, 'total': students.count(),
    })


# ============================================================
# ПРОФИЛЬ
# ============================================================

@login_required
def edit_profile(request):
    from django.contrib.auth import update_session_auth_hash

    user = request.user
    profile = user.profile
    success = False
    error = None

    school = profile.school
    all_subjects = Subject.objects.filter(is_active=True)
    profile_subjects = Subject.objects.filter(category='profile', is_active=True)
    all_classes = school.classes.all() if school else []

    if profile.role == 'teacher':
        occupied_ids = UserProfile.objects.filter(
            school=school, is_homeroom_teacher=True
        ).exclude(user=user).values_list('homeroom_class_id', flat=True)
        all_classes = all_classes.exclude(id__in=occupied_ids)

    selected_subject_ids = list(profile.subjects.values_list('id', flat=True))

    if request.method == 'POST':
        first_name = request.POST.get('first_name', '').strip()
        last_name = request.POST.get('last_name', '').strip()
        old_password = request.POST.get('old_password', '')
        new_password = request.POST.get('new_password', '')
        new_password2 = request.POST.get('new_password2', '')

        if first_name:
            user.first_name = first_name
        if last_name:
            user.last_name = last_name

        if new_password or new_password2 or old_password:
            if not user.check_password(old_password):
                error = _("Старый пароль неверный.")
            elif new_password != new_password2:
                error = _("Новые пароли не совпадают.")
            elif len(new_password) < 6:
                error = _("Пароль минимум 6 символов.")
            else:
                user.set_password(new_password)

        if profile.role == 'student' and not error:
            class_id = request.POST.get('school_class')
            subj1_id = request.POST.get('profile_subject_1')
            subj2_id = request.POST.get('profile_subject_2')
            if not class_id:
                error = _("Выберите класс.")
            elif not subj1_id or not subj2_id:
                error = _("Выберите два профильных предмета.")
            elif subj1_id == subj2_id:
                error = _("Профильные предметы должны быть разными.")
            else:
                try:
                    school_class = SchoolClass.objects.get(id=class_id, school=school)
                    subj1 = Subject.objects.get(id=subj1_id)
                    subj2 = Subject.objects.get(id=subj2_id)
                    profile.school_class = school_class
                    profile.profile_subject_1 = subj1
                    profile.profile_subject_2 = subj2
                except (SchoolClass.DoesNotExist, Subject.DoesNotExist):
                    error = _("Ошибка при сохранении.")

        if profile.role == 'teacher' and not error:
            is_subject = request.POST.get('is_subject_teacher') == 'on'
            is_homeroom = request.POST.get('is_homeroom_teacher') == 'on'
            if not is_subject and not is_homeroom:
                error = _("Выберите хотя бы одно.")
            else:
                if is_subject:
                    subject_ids = request.POST.getlist('subjects')
                    if not subject_ids:
                        error = _("Выберите хотя бы один предмет.")
                    else:
                        profile.subjects.set(Subject.objects.filter(id__in=subject_ids))
                if is_homeroom and not error:
                    homeroom_id = request.POST.get('homeroom_class')
                    if not homeroom_id:
                        error = _("Выберите класс.")
                    else:
                        try:
                            target_class = SchoolClass.objects.get(id=homeroom_id, school=school)
                            existing = UserProfile.objects.filter(
                                homeroom_class=target_class,
                                is_homeroom_teacher=True,
                            ).exclude(user=user).exists()
                            if existing:
                                error = _("У класса %(cls)s уже есть классный руководитель.") % {'cls': target_class}
                            else:
                                profile.homeroom_class = target_class
                        except SchoolClass.DoesNotExist:
                            error = _("Класс не найден.")
                if not error:
                    profile.is_subject_teacher = is_subject
                    profile.is_homeroom_teacher = is_homeroom
                    if not is_homeroom:
                        profile.homeroom_class = None

        if not error:
            user.save()
            profile.save()
            success = True
            if new_password:
                update_session_auth_hash(request, user)

    return render(request, 'core/edit_profile.html', {
        'profile': profile, 'success': success, 'error': error,
        'all_subjects': all_subjects, 'profile_subjects': profile_subjects,
        'all_classes': all_classes, 'selected_subject_ids': selected_subject_ids,
    })


# ============================================================
# СҰРАҚ ТОЛЫҚ КӨРІНІСІ
# ============================================================

@login_required
def question_detail(request, question_id):
    """Сұрақты толық көру. Рұқсат: автор, директор, завуч."""
    profile = request.user.profile
    try:
        q = Question.objects.select_related(
            'author', 'subject', 'school', 'group', 'approved_by'
        ).get(id=question_id)
    except Question.DoesNotExist:
        return redirect('core:home')

    can_view = False
    if profile.role in ['director', 'zavuch']:
        if q.school == profile.school:
            can_view = True
    elif profile.role == 'teacher':
        if q.author == request.user:
            can_view = True

    if not can_view:
        return render(request, 'core/access_denied.html', {
            'message': _("Сізде бұл сұрақты көруге құқық жоқ.")
        })

    matching_detail = []
    if q.question_type == 'matching' and q.matching_data:
        options = q.matching_data.get('options', [])
        subs = q.matching_data.get('sub_questions', [])
        for sub_idx, sub in enumerate(subs):
            matching_detail.append({
                'text': sub.get('text', ''),
                'correct': sub.get('correct', ''),
                'options': options,
            })

    group_info = None
    if q.group:
        group_info = {
            'title': q.group.title,
            'context_text': q.group.context_text,
            'difficulty': q.group.get_difficulty_display(),
            'current': q.group.current_question_count,
            'expected': q.group.expected_question_count,
        }

    siblings = []
    if q.group:
        siblings = q.group.questions.filter(
            is_deleted_by_author=False
        ).exclude(id=q.id).order_by('id')

    return render(request, 'core/question_detail.html', {
        'q': q,
        'profile': profile,
        'matching_detail': matching_detail,
        'group_info': group_info,
        'siblings': siblings,
        'back_url': request.META.get('HTTP_REFERER', reverse('core:home')),
    })


# ============================================================
# КОНТЕКСТ ТОЛЫҚ КӨРІНІСІ
# ============================================================

@login_required
def group_detail(request, group_id):
    """Контексті толық көру (мәтін + барлық сұрақтары).
    Рұқсат: автор, директор, завуч.
    """
    profile = request.user.profile
    try:
        g = QuestionGroup.objects.select_related(
            'author', 'subject', 'school', 'approved_by'
        ).prefetch_related('questions__author').get(id=group_id)
    except QuestionGroup.DoesNotExist:
        return redirect('core:home')

    can_view = False
    if profile.role in ['director', 'zavuch']:
        if g.school == profile.school:
            can_view = True
    elif profile.role == 'teacher':
        if g.author == request.user:
            can_view = True

    if not can_view:
        return render(request, 'core/access_denied.html', {
            'message': _("Сізде бұл контексті көруге құқық жоқ.")
        })

    questions_data = []
    for q in g.questions.filter(is_deleted_by_author=False).order_by('id'):
        matching_detail = []
        if q.question_type == 'matching' and q.matching_data:
            options = q.matching_data.get('options', [])
            for sub_idx, sub in enumerate(q.matching_data.get('sub_questions', [])):
                matching_detail.append({
                    'text': sub.get('text', ''),
                    'correct': sub.get('correct', ''),
                    'options': options,
                })
        questions_data.append({
            'q': q,
            'matching_detail': matching_detail,
        })

    return render(request, 'core/group_detail.html', {
        'g': g,
        'profile': profile,
        'questions_data': questions_data,
        'back_url': request.META.get('HTTP_REFERER', reverse('core:home')),
    })