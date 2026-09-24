from django.db import models
from django.contrib.auth.models import User
from django.core.validators import MinValueValidator, MaxValueValidator
import secrets


# ============================================================
# ТҰРАҚТЫЛАР
# ============================================================

DIFFICULTY_A = 'A'
DIFFICULTY_B = 'B'
DIFFICULTY_C = 'C'

DIFFICULTY_CHOICES = [
    (DIFFICULTY_A, 'Оңай'),
    (DIFFICULTY_B, 'Орташа'),
    (DIFFICULTY_C, 'Қиын'),
]

CONTEXT_DIFFICULTY_COUNT = {
    DIFFICULTY_A: 5,
    DIFFICULTY_B: 3,
    DIFFICULTY_C: 2,
}

TEST_DISTRIBUTION = {
    'kaz_history':   {DIFFICULTY_A: 10, DIFFICULTY_B: 6,  DIFFICULTY_C: 4},
    'reading':       {DIFFICULTY_A: 5,  DIFFICULTY_B: 3,  DIFFICULTY_C: 2},
    'math_literacy': {DIFFICULTY_A: 5,  DIFFICULTY_B: 3,  DIFFICULTY_C: 2},
    'profile':       {DIFFICULTY_A: 20, DIFFICULTY_B: 12, DIFFICULTY_C: 8},
}


# ============================================================
# SUBJECT
# ============================================================

class Subject(models.Model):
    CATEGORY_CHOICES = [
        ('mandatory', 'Обязательный'),
        ('profile', 'Профильный'),
    ]
    CODE_CHOICES = [
        ('kaz_history', 'История Казахстана'),
        ('reading', 'Грамотность чтения'),
        ('math_literacy', 'Математическая грамотность'),
        ('profile', 'Профильный предмет'),
    ]

    code = models.CharField(max_length=20, choices=CODE_CHOICES)
    name_kk = models.CharField(max_length=100, verbose_name="Название (каз)")
    name_ru = models.CharField(max_length=100, verbose_name="Название (рус)")
    name_en = models.CharField(max_length=100, blank=True, verbose_name="Название (англ)")
    category = models.CharField(max_length=20, choices=CATEGORY_CHOICES)
    is_active = models.BooleanField(default=True)

    def __str__(self):
        return self.name_ru

    @property
    def name(self):
        from django.utils.translation import get_language
        lang = get_language()
        if lang == 'kk' and self.name_kk:
            return self.name_kk
        if lang == 'en' and self.name_en:
            return self.name_en
        return self.name_ru


# ============================================================
# SCHOOL
# ============================================================

class School(models.Model):
    name = models.CharField(max_length=200)
    city = models.CharField(max_length=100)
    teacher_code = models.CharField(max_length=20, unique=True, blank=True)
    zavuch_code = models.CharField(max_length=20, unique=True, blank=True)
    student_code = models.CharField(max_length=20, unique=True, blank=True)
    director = models.OneToOneField(User, on_delete=models.SET_NULL, null=True,
                                    blank=True, related_name='directed_school')
    created_at = models.DateTimeField(auto_now_add=True)

    def save(self, *args, **kwargs):
        if not self.teacher_code:
            self.teacher_code = 'T-' + secrets.token_hex(4).upper()
        if not self.zavuch_code:
            self.zavuch_code = 'Z-' + secrets.token_hex(4).upper()
        if not self.student_code:
            self.student_code = 'S-' + secrets.token_hex(4).upper()
        super().save(*args, **kwargs)

    def __str__(self):
        return f"{self.name} ({self.city})"


class SchoolClass(models.Model):
    LETTER_CHOICES = [
        ('А', 'А'), ('Б', 'Б'), ('В', 'В'),
        ('Г', 'Г'), ('Д', 'Д'), ('Е', 'Е'),
    ]
    school = models.ForeignKey(School, on_delete=models.CASCADE, related_name='classes')
    grade = models.IntegerField(validators=[MinValueValidator(9), MaxValueValidator(11)])
    letter = models.CharField(max_length=2, choices=LETTER_CHOICES)

    class Meta:
        unique_together = ('school', 'grade', 'letter')
        ordering = ['grade', 'letter']

    def __str__(self):
        return f"{self.grade}{self.letter}"


# ============================================================
# USER PROFILE
# ============================================================

class UserProfile(models.Model):
    ROLE_CHOICES = [
        ('student', 'Ученик'),
        ('teacher', 'Учитель'),
        ('director', 'Директор'),
        ('zavuch', 'Завуч'),
    ]

    user = models.OneToOneField(User, on_delete=models.CASCADE, related_name='profile')
    school = models.ForeignKey(School, on_delete=models.CASCADE, null=True, blank=True)
    role = models.CharField(max_length=20, choices=ROLE_CHOICES)

    is_subject_teacher = models.BooleanField(default=False)
    subjects = models.ManyToManyField(Subject, blank=True, related_name='teachers')

    is_homeroom_teacher = models.BooleanField(default=False)
    homeroom_class = models.ForeignKey(SchoolClass, on_delete=models.SET_NULL,
                                       null=True, blank=True, related_name='homeroom_teacher')

    school_class = models.ForeignKey(SchoolClass, on_delete=models.SET_NULL,
                                     null=True, blank=True, related_name='students')
    profile_subject_1 = models.ForeignKey(Subject, on_delete=models.SET_NULL, null=True,
                                          blank=True, related_name='students_profile_1')
    profile_subject_2 = models.ForeignKey(Subject, on_delete=models.SET_NULL, null=True,
                                          blank=True, related_name='students_profile_2')

    class Meta:
        constraints = [
            models.UniqueConstraint(
                fields=['homeroom_class'],
                condition=models.Q(is_homeroom_teacher=True),
                name='unique_homeroom_teacher_per_class',
            ),
        ]

    def __str__(self):
        return f"{self.user.get_full_name()} — {self.get_role_display()}"


# ============================================================
# QUESTION GROUP
# ============================================================

class QuestionGroup(models.Model):
    BLOCK_CHOICES = [
        ('reading', 'Грамотность чтения'),
        ('kaz_history', 'История Казахстана'),
        ('math_literacy', 'Математическая грамотность'),
        ('profile', 'Профильный предмет'),
    ]

    school = models.ForeignKey(School, on_delete=models.CASCADE, related_name='question_groups')
    title = models.CharField(max_length=200)
    context_text = models.TextField()
    block = models.CharField(max_length=20, choices=BLOCK_CHOICES)
    subject = models.ForeignKey(Subject, on_delete=models.SET_NULL, null=True,
                                blank=True, related_name='question_groups')
    author = models.ForeignKey(User, on_delete=models.SET_NULL, null=True,
                               blank=True, related_name='authored_groups')
    status = models.CharField(max_length=20, choices=[
        ('draft', 'Черновик'),
        ('pending', 'На проверке'),
        ('approved', 'Одобрен'),
        ('rejected', 'Отклонён'),
    ], default='draft')

    difficulty = models.CharField(
        max_length=1, choices=DIFFICULTY_CHOICES, default=DIFFICULTY_A,
    )

    approved_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True,
                                    blank=True, related_name='approved_groups')
    approved_at = models.DateTimeField(null=True, blank=True)
    created_at = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return f"{self.title} ({self.get_difficulty_display()})"

    @property
    def expected_question_count(self):
        # Профильде ӘРҚАШАН 5 сұрақ
        if self.block == 'profile':
            return 5
        return CONTEXT_DIFFICULTY_COUNT.get(self.difficulty, 0)

    @property
    def current_question_count(self):
        return self.questions.filter(is_deleted_by_author=False).count()

    @property
    def is_complete(self):
        return self.current_question_count == self.expected_question_count


# ============================================================
# QUESTION
# ============================================================

class Question(models.Model):
    QUESTION_TYPE = [
        ('single', 'Один правильный ответ'),
        ('multiple', 'Несколько правильных'),
        ('matching', 'Сопоставление'),
    ]
    KIND_CHOICES = [
        ('standard', 'Стандартный'),
        ('context', 'Контекстный'),
    ]
    STATUS_CHOICES = [
        ('draft', 'Черновик'),
        ('pending', 'На проверке'),
        ('approved', 'Одобрен'),
        ('rejected', 'Отклонён'),
    ]
    BLOCK_CHOICES = [
        ('kaz_history', 'История Казахстана'),
        ('reading', 'Грамотность чтения'),
        ('math_literacy', 'Математическая грамотность'),
        ('profile', 'Профильный предмет'),
    ]

    block = models.CharField(max_length=20, choices=BLOCK_CHOICES)
    subject = models.ForeignKey(Subject, on_delete=models.SET_NULL, null=True,
                                blank=True, related_name='questions')
    group = models.ForeignKey(QuestionGroup, on_delete=models.SET_NULL, null=True,
                              blank=True, related_name='questions')
    question_type = models.CharField(max_length=10, choices=QUESTION_TYPE, default='single')
    kind = models.CharField(max_length=20, choices=KIND_CHOICES, default='standard')

    difficulty = models.CharField(
        max_length=1, choices=DIFFICULTY_CHOICES, default=DIFFICULTY_A,
        db_index=True,
    )

    text = models.TextField()
    option_a = models.CharField(max_length=500, blank=True)
    option_b = models.CharField(max_length=500, blank=True)
    option_c = models.CharField(max_length=500, blank=True)
    option_d = models.CharField(max_length=500, blank=True)
    option_e = models.CharField(max_length=500, blank=True)
    option_f = models.CharField(max_length=500, blank=True)
    correct_answer = models.CharField(max_length=6, blank=True)

    # ✅ Бір matching сұрағының деректері:
    # {"options": ["a","b","c","d"], "sub_questions": [{"text":"...","correct":"a"}, ...]}
    matching_data = models.JSONField(null=True, blank=True)

    image = models.ImageField(
        upload_to='questions/images/%Y/%m/', null=True, blank=True,
        verbose_name="Сурет",
    )

    author = models.ForeignKey(User, on_delete=models.SET_NULL, null=True,
                               blank=True, related_name='authored_questions')
    school = models.ForeignKey(School, on_delete=models.CASCADE, null=True,
                               blank=True, related_name='questions')
    status = models.CharField(max_length=20, choices=STATUS_CHOICES, default='draft')

    is_deleted_by_author = models.BooleanField(default=False)

    approved_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True,
                                    blank=True, related_name='approved_questions')
    approved_at = models.DateTimeField(null=True, blank=True)
    times_used = models.IntegerField(default=0)
    is_locked = models.BooleanField(default=False)
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ['-created_at']
        indexes = [
            models.Index(fields=['block', 'difficulty', 'status']),
            models.Index(fields=['subject', 'difficulty', 'status']),
        ]

    def __str__(self):
        subj = self.subject.name if self.subject else self.get_block_display()
        return f"[{subj}/{self.get_difficulty_display()}] {self.text[:50]}"

    @property
    def max_score(self):
        if self.question_type == 'multiple':
            return 2
        if self.question_type == 'matching':
            return 2  # бір matching сұрағы = 2 балл макс
        return 1


# ============================================================
# EXAM SESSION
# ============================================================

class ExamSession(models.Model):
    school = models.ForeignKey(School, on_delete=models.CASCADE, related_name='exam_sessions')
    title = models.CharField(max_length=200)
    opens_at = models.DateTimeField()
    closes_at = models.DateTimeField()
    created_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True,
                                   blank=True, related_name='created_sessions')
    created_at = models.DateTimeField(auto_now_add=True)
    classes = models.ManyToManyField(SchoolClass, blank=True, related_name='exam_sessions')
    is_active = models.BooleanField(default=True)

    def __str__(self):
        return f"{self.title} ({self.opens_at:%d.%m.%Y} — {self.closes_at:%d.%m.%Y})"

    @property
    def is_open_now(self):
        from django.utils import timezone
        now = timezone.now()
        return self.is_active and self.opens_at <= now <= self.closes_at


# ============================================================
# TEST ATTEMPT
# ============================================================

class TestAttempt(models.Model):
    STATUS_CHOICES = [
        ('in_progress', 'В процессе'),
        ('finished', 'Завершён'),
        ('annulled', 'Аннулирован'),
    ]

    student = models.ForeignKey(User, on_delete=models.CASCADE, related_name='test_attempts')
    session = models.ForeignKey(ExamSession, on_delete=models.SET_NULL, null=True,
                                blank=True, related_name='attempts')
    started_at = models.DateTimeField(auto_now_add=True)
    finished_at = models.DateTimeField(null=True, blank=True)

    tab_switch_count = models.IntegerField(default=0)
    status = models.CharField(max_length=20, choices=STATUS_CHOICES, default='in_progress')

    total_score = models.IntegerField(default=0)
    score_kaz_history = models.IntegerField(default=0)
    score_reading = models.IntegerField(default=0)
    score_math_literacy = models.IntegerField(default=0)
    score_profile_1 = models.IntegerField(default=0)
    score_profile_2 = models.IntegerField(default=0)

    passed_thresholds = models.BooleanField(default=False)
    profile_subject_1 = models.CharField(max_length=100, blank=True)
    profile_subject_2 = models.CharField(max_length=100, blank=True)

    def __str__(self):
        return f"{self.student.get_full_name()} — {self.get_status_display()} ({self.total_score})"


# ============================================================
# ANSWER
# ============================================================

class Answer(models.Model):
    attempt = models.ForeignKey(TestAttempt, on_delete=models.CASCADE, related_name='answers')
    question = models.ForeignKey(Question, on_delete=models.SET_NULL, null=True, blank=True)

    question_snapshot = models.TextField(blank=True)
    correct_answer_snapshot = models.CharField(max_length=6, blank=True)
    question_type_snapshot = models.CharField(max_length=10, blank=True)
    subject_snapshot = models.CharField(max_length=30, blank=True)
    difficulty_snapshot = models.CharField(max_length=1, blank=True)
    kind_snapshot = models.CharField(max_length=20, blank=True)
    matching_data_snapshot = models.JSONField(null=True, blank=True)
    image_snapshot = models.CharField(max_length=500, blank=True)

    chosen = models.CharField(max_length=6, blank=True)
    # Matching үшін: {"0": "2", "1": "2"} — sub_question индексі → таңдаған нұсқа
    matching_answers = models.JSONField(null=True, blank=True)

    score = models.IntegerField(default=0)
    group_context_snapshot = models.TextField(blank=True)
    group_title_snapshot = models.CharField(max_length=200, blank=True)

    class Meta:
        unique_together = ('attempt', 'question')

    def __str__(self):
        return f"Ответ на {self.question_id}: {self.chosen or 'matching'} → {self.score} б."