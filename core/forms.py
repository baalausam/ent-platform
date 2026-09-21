from django import forms
from django.contrib.auth.models import User
from django.utils.translation import gettext_lazy as _

from .models import (School, SchoolClass, Subject, Question,
                     ExamSession, QuestionGroup)


# ============================================================
# РЕГИСТРАЦИЯ ДИРЕКТОРА
# ============================================================

class DirectorRegistrationForm(forms.Form):
    first_name = forms.CharField(max_length=100, label=_("Имя"))
    last_name = forms.CharField(max_length=100, label=_("Фамилия"))
    username = forms.CharField(max_length=150, label=_("Логин"))
    password = forms.CharField(widget=forms.PasswordInput, label=_("Пароль"))
    password2 = forms.CharField(widget=forms.PasswordInput, label=_("Повторите пароль"))

    school_name = forms.CharField(max_length=200, label=_("Название школы"))
    city = forms.CharField(max_length=100, label=_("Город"))

    def clean_username(self):
        username = self.cleaned_data['username']
        if User.objects.filter(username=username).exists():
            raise forms.ValidationError(_("Такой логин уже занят."))
        return username

    def clean(self):
        cleaned = super().clean()
        p1 = cleaned.get('password')
        p2 = cleaned.get('password2')
        if p1 and p2 and p1 != p2:
            raise forms.ValidationError(_("Пароли не совпадают."))
        return cleaned

    def clean_school_name(self):
        name = self.cleaned_data['school_name']
        if School.objects.filter(name=name).exists():
            raise forms.ValidationError(
                _("Школа с таким названием уже зарегистрирована. "
                  "Если вы её директор — обратитесь к администратору.")
            )
        return name


# ============================================================
# РЕГИСТРАЦИЯ УЧИТЕЛЯ / ЗАВУЧА
# ============================================================

class TeacherRegistrationForm(forms.Form):
    first_name = forms.CharField(max_length=100, label=_("Имя"))
    last_name = forms.CharField(max_length=100, label=_("Фамилия"))
    username = forms.CharField(max_length=150, label=_("Логин"))
    password = forms.CharField(widget=forms.PasswordInput, label=_("Пароль"))
    password2 = forms.CharField(widget=forms.PasswordInput, label=_("Повторите пароль"))
    teacher_code = forms.CharField(max_length=20, label=_("Код школы"))

    def clean_username(self):
        username = self.cleaned_data['username']
        if User.objects.filter(username=username).exists():
            raise forms.ValidationError(_("Такой логин уже занят."))
        return username

    def clean(self):
        cleaned = super().clean()
        p1 = cleaned.get('password')
        p2 = cleaned.get('password2')
        if p1 and p2 and p1 != p2:
            raise forms.ValidationError(_("Пароли не совпадают."))
        return cleaned


# ============================================================
# РЕГИСТРАЦИЯ УЧЕНИКА
# ============================================================

class StudentRegistrationForm(forms.Form):
    first_name = forms.CharField(max_length=100, label=_("Имя"))
    last_name = forms.CharField(max_length=100, label=_("Фамилия"))
    username = forms.CharField(max_length=150, label=_("Логин"))
    password = forms.CharField(widget=forms.PasswordInput, label=_("Пароль"))
    password2 = forms.CharField(widget=forms.PasswordInput, label=_("Повторите пароль"))
    student_code = forms.CharField(max_length=20, label=_("Код школы (для учеников)"))

    def clean_username(self):
        username = self.cleaned_data['username']
        if User.objects.filter(username=username).exists():
            raise forms.ValidationError(_("Такой логин уже занят."))
        return username

    def clean(self):
        cleaned = super().clean()
        p1 = cleaned.get('password')
        p2 = cleaned.get('password2')
        if p1 and p2 and p1 != p2:
            raise forms.ValidationError(_("Пароли не совпадают."))
        return cleaned


# ============================================================
# ВОПРОС
# ============================================================

class QuestionForm(forms.Form):
    BLOCK_CHOICES = [
        ('kaz_history', _('История Казахстана')),
        ('reading', _('Грамотность чтения')),
        ('math_literacy', _('Математическая грамотность')),
        ('profile', _('Профильный предмет')),
    ]

    QUESTION_TYPE_CHOICES = [
        ('single', _('Один правильный ответ')),
        ('multiple', _('Несколько правильных ответов')),
    ]

    block = forms.ChoiceField(choices=BLOCK_CHOICES, label=_("Блок ЕНТ"))
    subject = forms.ModelChoiceField(
        queryset=Subject.objects.none(),
        required=False,
        label=_("Профильный предмет"),
    )
    question_type = forms.ChoiceField(
        choices=QUESTION_TYPE_CHOICES,
        label=_("Тип вопроса"),
    )
    text = forms.CharField(
        widget=forms.Textarea(attrs={'rows': 4}),
        label=_("Текст вопроса"),
    )
    option_a = forms.CharField(max_length=500, label=_("Вариант A"))
    option_b = forms.CharField(max_length=500, label=_("Вариант B"))
    option_c = forms.CharField(max_length=500, label=_("Вариант C"))
    option_d = forms.CharField(max_length=500, label=_("Вариант D"))
    option_e = forms.CharField(max_length=500, required=False, label=_("Вариант E"))
    option_f = forms.CharField(max_length=500, required=False, label=_("Вариант F"))
    correct_answer = forms.CharField(
        max_length=6,
        label=_("Правильный ответ"),
        help_text=_("Для одного: A, B, C или D. "
                    "Для нескольких: AB, ACD, ABCD, ABCDEF."),
    )

    def __init__(self, *args, user=None, **kwargs):
        super().__init__(*args, **kwargs)
        if user:
            profile = user.profile
            self.fields['subject'].queryset = profile.subjects.filter(category='profile')

    def clean(self):
        cleaned = super().clean()
        block = cleaned.get('block')
        subject = cleaned.get('subject')
        qtype = cleaned.get('question_type')
        correct = (cleaned.get('correct_answer') or '').strip().upper()

        if block == 'profile' and not subject:
            self.add_error('subject', _("Выберите профильный предмет."))

        valid_letters = set('ABCDEF')
        if not correct:
            self.add_error('correct_answer', _("Укажите правильный ответ."))
        elif not set(correct).issubset(valid_letters):
            self.add_error('correct_answer', _("Допустимы только A, B, C, D, E, F."))
        elif qtype == 'single' and len(correct) != 1:
            self.add_error('correct_answer', _("Для одного ответа укажите одну букву."))
        elif qtype == 'multiple' and len(correct) < 2:
            self.add_error('correct_answer', _("Для нескольких ответов укажите минимум две буквы."))
        elif len(correct) != len(set(correct)):
            self.add_error('correct_answer', _("Буквы не должны повторяться."))

        return cleaned


# ============================================================
# КОНТЕКСТ
# ============================================================

class QuestionGroupForm(forms.ModelForm):
    class Meta:
        model = QuestionGroup
        fields = ['title', 'context_text']
        widgets = {
            'context_text': forms.Textarea(attrs={'rows': 12}),
        }


# ============================================================
# СЕССИЯ ЕНТ
# ============================================================

class ExamSessionForm(forms.Form):
    title = forms.CharField(max_length=200, label=_("Название сессии"))
    opens_at = forms.DateTimeField(
        label=_("Начало"),
        widget=forms.DateTimeInput(attrs={'type': 'datetime-local'}),
    )
    closes_at = forms.DateTimeField(
        label=_("Окончание"),
        widget=forms.DateTimeInput(attrs={'type': 'datetime-local'}),
    )
    classes = forms.ModelMultipleChoiceField(
        queryset=SchoolClass.objects.none(),
        required=False,
        widget=forms.CheckboxSelectMultiple,
        label=_("Классы (пусто = для всех)"),
    )

    def __init__(self, *args, school=None, **kwargs):
        super().__init__(*args, **kwargs)
        if school:
            self.fields['classes'].queryset = school.classes.all()

    def clean(self):
        cleaned = super().clean()
        opens = cleaned.get('opens_at')
        closes = cleaned.get('closes_at')
        if opens and closes and opens >= closes:
            raise forms.ValidationError(_("Окончание должно быть позже начала."))
        return cleaned