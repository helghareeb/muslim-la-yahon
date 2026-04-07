# المسلم لا يهون على الله وإن هان على الناس

هذا المستودع يحتوي على المادة المصدرية للكتاب بصيغتي `docx` و`pptx`، مع نسخة `LaTeX` عربية جاهزة للنشر والطباعة.

## الملفات المهمة

- `build_book.py`: مولد الكتاب من ملف الـ`Word`.
- `main.tex`: النسخة النهائية المنسقة بـ`XeLaTeX`.
- `main.pdf`: الخرج المبني.
- `assets/ppt-cover-bg.jpg`: الخلفية المستخرجة من العرض التقديمي لغلاف الكتاب.
- `extracted-source.txt`: نسخة نصية مرجعية للمراجعة والمقابلة مع الملف الأصلي.
- `PROJECT_STATUS.md`: حالة المشروع الحالية وآخر نقطة وصلنا إليها.
- `SESSION_LOG.md`: سجل مختصر بما تم إنجازه في كل جلسة.
- `NEXT_STEPS.md`: الخطوات التالية المقترحة لاستكمال العمل.

## إعادة البناء

```bash
python3 build_book.py
xelatex -interaction=nonstopmode -halt-on-error main.tex
xelatex -interaction=nonstopmode -halt-on-error main.tex
```
