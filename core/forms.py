from django import forms
from django.contrib.auth.models import User
from django.utils.translation import gettext_lazy as _

from .models import (
    School, SchoolClass, Subject, Question, ExamSession,
    QuestionGroup, DIFFICULTY_CHOICES,
)


# ============================================================
# ТІРКЕЛУ
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
            raise forms.ValidationError(_("Школа с таким названием уже зарегистрирована."))
        return name


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


class StudentRegistrationForm(forms.Form):
    first_name = forms.CharField(max_length=100, label=_("Имя"))
    last_name = forms.CharField(max_length=100, label=_("Фамилия"))
    username = forms.CharField(max_length=150, label=_("Логин"))
    password = forms.CharField(widget=forms.PasswordInput, label=_("Пароль"))
    password2 = forms.CharField(widget=forms.PasswordInput, label=_("Повторите пароль"))
    student_code = forms.CharField(max_length=20, label=_("Код школы"))

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
# СҰРАҚ
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
        ('matching', _('Сопоставление')),
    ]
    DIFFICULTY_CHOICES = [
        ('A', _('Оңай')),
        ('B', _('Орташа')),
        ('C', _('Қиын')),
    ]

    block = forms.ChoiceField(choices=BLOCK_CHOICES, label=_("Блок ЕНТ"))
    subject = forms.ModelChoiceField(
        queryset=Subject.objects.none(), required=False,
        label=_("Профильный предмет"),
    )
    difficulty = forms.ChoiceField(
        choices=DIFFICULTY_CHOICES, label=_("Қиындық деңгейі"),
        widget=forms.RadioSelect,
    )
    question_type = forms.ChoiceField(
        choices=QUESTION_TYPE_CHOICES, label=_("Тип вопроса"),
    )

    # ✅ MathLive формуласы
    text = forms.CharField(
        widget=forms.Textarea(attrs={'rows': 4}),
        label=_("Формула (LaTeX)"),
        required=False,
    )

    # ✅ Қарапайым мәтін
    text_plain = forms.CharField(
        widget=forms.Textarea(attrs={'rows': 4}),
        label=_("Қарапайым мәтін"),
        required=False,
    )

    option_a = forms.CharField(max_length=500, required=False, label=_("Вариант A"))
    option_b = forms.CharField(max_length=500, required=False, label=_("Вариант B"))
    option_c = forms.CharField(max_length=500, required=False, label=_("Вариант C"))
    option_d = forms.CharField(max_length=500, required=False, label=_("Вариант D"))
    option_e = forms.CharField(max_length=500, required=False, label=_("Вариант E"))
    option_f = forms.CharField(max_length=500, required=False, label=_("Вариант F"))
    correct_answer = forms.CharField(max_length=6, required=False, label=_("Правильный ответ"))

    image = forms.ImageField(
        required=False, label=_("Сурет (міндетті емес)"),
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
        text = (cleaned.get('text') or '').strip()
        text_plain = (cleaned.get('text_plain') or '').strip()

        # ✅ Кем дегенде біреуі толтырылуы керек
        if not text and not text_plain:
            self.add_error('text_plain', _("Мәтін енгізіңіз (екі өрістің біреуі)"))

        if block == 'profile' and not subject:
            self.add_error('subject', _("Выберите профильный предмет."))

        if qtype == 'matching':
            return cleaned

        valid_letters = set('ABCDEF')
        if not correct:
            self.add_error('correct_answer', _("Укажите правильный ответ."))
        elif not set(correct).issubset(valid_letters):
            self.add_error('correct_answer', _("Допустимы только A-F."))
        elif qtype == 'single' and len(correct) != 1:
            self.add_error('correct_answer', _("Для одного ответа — одна буква."))
        elif qtype == 'multiple' and len(correct) < 2:
            self.add_error('correct_answer', _("Минимум две буквы."))

        return cleaned


# ============================================================
# КОНТЕКСТ
# ============================================================

class QuestionGroupForm(forms.ModelForm):
    difficulty = forms.ChoiceField(
        choices=[
            ('A', _('Оңай — 5 сұрақ')),
            ('B', _('Орташа — 3 сұрақ')),
            ('C', _('Қиын — 2 сұрақ')),
        ],
        label=_("Контекст деңгейі (тек оқу сауаттылығы үшін)"),
        widget=forms.RadioSelect,
        required=False,
    )

    image = forms.ImageField(
        required=False,
        label=_("Контекст суреті (міндетті емес)"),
        help_text=_("JPG, PNG, GIF. Максимум 5 МБ."),
    )

    class Meta:
        model = QuestionGroup
        fields = ['title', 'context_text', 'difficulty', 'image']
        widgets = {
            'context_text': forms.Textarea(attrs={'rows': 12}),
        }


# ============================================================
# СЕССИЯ
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
        queryset=SchoolClass.objects.none(), required=False,
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