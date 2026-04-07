#!/usr/bin/env python3
from __future__ import annotations

import re
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path


ROOT = Path(__file__).resolve().parent
DOCX_PATH = ROOT / "بحث - المسلم لا يهون على الله وإن هان على الناس.docx"
OUTPUT_PATH = ROOT / "main.tex"

NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
SYMS = {
    "F047": "ﷺ",
    "F062": "ﷻ",
    "F069": "رضي الله عنها",
    "F072": "رحمه الله تعالى",
    "F068": "رضي الله عنه",
    "F06A": "رضي الله عنهم",
    "F06E": "عليه السلام",
}

PARA_OVERRIDES = {
    3: "الحمد لله الذي أعزَّ أهلَ الإيمان بطاعته، ورفع قدرَهم بالاستمساك بدينه، وجعل لهم من نوره وهدايته ما يثبتهم عند الفتن والشدائد. والصلاة والسلام على نبينا محمد ﷺ، إمام الصابرين، وقدوة الثابتين، وعلى آله وصحبه أجمعين. أما بعد؛ فإن ما يلقاه المسلمون في أزمنة الضعف والاستضعاف من صور الإيذاء والإذلال قد يوقع في بعض النفوس سؤالًا مؤلمًا: هل ما أصابهم دليلُ هوانهم على ربهم؟ فجاء هذا البحث ليدفع هذا الوهم من أصله، ويقرر أنَّ المسلم لا يهون على الله ﷻ وإن اشتدَّ عليه البلاء، أو خانه الناس، أو تقلَّبت به الأحوال. وسيعتمد هذا البحث على نصوص الكتاب والسنة، وعلى ما فيها من قصص الأنبياء، وسير الصحابة، وكلام أهل العلم، لبيان أنَّ الابتلاء سُنَّةٌ ربانية، وأنَّ الكرامة الحقة ليست في زخارف الدنيا وأسباب التمكين الظاهرة، وإنما في ثبات القلب على الإيمان، ولزوم الصراط المستقيم، وحسن الصلة برب العالمين. ثم يختم بوصايا جامعة تعين على الصبر، وتبعث على العزّة بالإيمان، وتردُّ القلب إلى منابع الثبات واليقين.",
    4: "من هدي الأنبياء في الصبر والثبات",
    5: "نبينا محمد ﷺ إمام الصابرين",
    6: "تجلَّى صبر النبي ﷺ على الابتلاء في أكمل صوره؛ فما صدَّته شدةُ الفقر والجوع عن دعوته، ولا أوهنت عزيمتَه قسوةُ الأذى، بل مضى قائمًا بأمر ربِّه، محتسبًا، صابرًا، مربّيًا لأصحابه على الثبات واليقين. وقد لقي ﷺ من قومه التكذيبَ والسخريةَ والأذى، وفارق من أحبَّ من عمٍّ وزوجةٍ وأولاد، ومع ذلك لم يعرف قلبُه جزعًا ولا لسانُه تسخُّطًا، بل كان رضاه بربه أعظم، وصبرُه على أمره أكمل. وفي سيرته ﷺ أعظم شاهدٍ على أنَّ البلاء لا يقدح في رفعة العبد عند الله، بل قد يكون ميدانًا يظهر فيه صدق العبودية، وكمال الصبر، وعلوُّ المنزلة.",
    68: "في سِيَر الصحابة أسوةٌ بعد الأنبياء",
    69: "لقد ورث صحابة رسول الله ﷺ عن نبيهم معاني الصبر والثبات واليقين، فثبتوا على الحق مع ما نالهم من تعذيبٍ وإقصاءٍ وفقرٍ ومفارقةٍ للأهل والأوطان. ولم تكن شدائد الطريق سببًا في ضعف إيمانهم، بل زادتهم تمسكًا بالدين، وبذلًا في سبيل الله، ورضًا بما عنده، حتى صاروا للأمة من بعدهم أئمةً في الصبر، ومناراتٍ في احتمال الأذى والثبات على المبدأ.",
    70: "مصعب بن عمير رضي الله عنه",
    71: "ومن أظهر النماذج في ذلك مصعب بن عمير رضي الله عنه؛ فقد كان من أنعم فتيان قريش عيشًا، ثم ترك نعيمها وزينتها حين عرف الحق، وآثر رضا الله ﷻ على ترف الدنيا ومتاعها. فانتقل من سعة العيش إلى ضيقه، ومن لين النشأة إلى خشونة الطريق، حتى قُتل شهيدًا ولم يوجد له من الكفن ما يستر جسده كله. وفي خبره شاهد بيّن على أنَّ قيمة العبد ليست بما يملك، وإنما بما يثبت عليه من الإيمان، وما يبذله لله من نفسٍ ومالٍ وراحة.",
    73: "الابتلاء سُنَّةٌ ماضية لا أمرٌ مفاجئ",
    85: "الابتلاء لا ينافي كرامة العبد على ربِّه",
    86: "لا تلازم بين التمكين الدنيوي وبين الكرامة عند الله ﷻ؛ فكم من مُنَعَّمٍ في الظاهر وهو عند الله خاسر، وكم من مبتلى ممتحَنٍ وهو عند الله كريم.",
    87: "وعليه؛ فما يَلحق المسلمين من المحن والشدائد لا يدلُّ على هوانهم على الله ﷻ، بل قد يكون من أعظم أبواب التمحيص والتزكية ورفع الدرجات.",
    90: "وكرامة المسلمين على الله ﷻ ثابتةٌ ما استمسكوا بكتاب الله تعالى، واعتصموا بسنة رسوله ﷺ، وصدقوا في الإيمان به والثبات على أمره.",
    91: "وفيما سبق من قصص الأنبياء والمرسلين، ثم من سِيَر الصحابة رضي الله عنهم، أعظمُ التعزية والسلوى. فليست الدنيا بموازينها الظاهرة هي معيارَ الكرامة والرفعة، بل المعيار الحق هو ما يثبّت الله به عبده على الإيمان، وما يرزقه من الصبر واليقين. وتأمل الآيات الآتية يتبين لك هذا المعنى جليًّا.",
    93: "وتأمل قول الله ﷻ في سورة الكهف:",
    95: "وتأمل أيضًا قول الله ﷻ:",
    98: "تمتيع الكفار في الدنيا ليس دليلَ كرامة",
    99: "والجواب عن هذا الإشكال يجلّيه قول الله ﷻ:",
    109: "شرح الصدر للإسلام من دلائل إرادة الخير",
    111: "من واجب المسلم عند الابتلاء",
    112: "الدعاء والافتقار إلى الله",
    117: "استحضار النهي عن الوهن والحزن",
    119: "الفرح بفضل الله ورحمته",
    124: "سبب شدَّة عداوتهم للإسلام والقرآن",
    126: "أسباب هذه العداوة",
    127: "لأنَّ الإسلام حقٌّ صافٍ، لا يقبل أن يُمازجَه باطل، ولا أن تستقر معه الأهواء والشهوات الفاسدة.",
    128: "لأنَّ القرآن الكريم يفضح الباطل، ويهتك دعاواه، ويقيم على الخلق حجةً ظاهرةً لا يملكون لها دفعًا.",
    129: "لأنهم عجزوا أن يأتوا بمثل هذا القرآن، مع قيام التحدي لهم، فكان عجزهم سببًا في ازدياد حَنَقهم وعداوتهم.",
    130: "لأنهم يعلمون أنَّ هذا الدين إذا خالطت بشاشتُه القلوبَ استعلى بها على الشهوات والمخاوف، فصار ذلك أدعى لغيظهم ومقاومتهم.",
    131: "لأنَّ الله ﷻ تكفَّل بحفظ كتابه، وجعل بقاءَ هذا الدين وظهورَ حجته من دلائل ربوبيته ووعده الصادق، قال تعالى: ﴿إِنَّ ٱللَّهَ لَا يُخۡلِفُ ٱلۡمِيعَادَ﴾ [آل عمران: 9]. فمن اعتصم بهذا الكتاب نال من أسباب العزّة والثبات بقدر تمسُّكه به.",
    133: "وخلاصة هذا البحث أنَّ المسلم له عند الله ﷻ قدرٌ عظيم، وأنَّ ما يناله من الابتلاء لا يدل على هوانه، بل قد يكون من دلائل العناية به، إذ يردُّه إلى ربِّه، ويطهِّر قلبه، ويُعلي درجته، ويُظهر صدقه وثباته.",
    134: "وقد دلَّت نصوص الكتاب والسنة، وسِيَر الأنبياء والصالحين، وكلام أهل العلم، على أنَّ طريق الحق لا ينفك عن الامتحان؛ غير أنَّ العبرة كلَّ العبرة بحسن الصبر، وصدق الالتجاء، ولزوم الوحي، والثبات على الجادة. فمن رزقه الله ذلك لم تضرَّه وحشة الطريق، ولا شماتة الأعداء، ولا تقلُّب الموازين في أعين الناس.",
    135: "ومن هنا كان لزامًا على المسلمين أن يتواصوا بالحق، ويتعاونوا على الصبر، وأن يُحيوا في أنفسهم معاني العزَّة بالإيمان، والثقة بوعد الرحمن، وحسن الظن برب العالمين. ومن أهم ما يعين على ذلك:",
    136: "التبصيرُ بمكانة المسلم عند ربِّه: وذلك بإحياء المعاني الشرعية التي تغرس في القلب أنَّ أهل الإيمان أهلُ ولايةٍ وكرامة، وأنَّ ما يعتريهم من ضعفٍ أو أذى لا يسلبهم هذه المنزلة ما داموا على الاستقامة.",
    137: "لزومُ كتاب الله وسنة رسوله ﷺ: فبهما حياة القلوب، وبهما النجاة من الفتن، وبهما يعرف المسلم مواضع العزِّ والخذلان، ويميز بين الموازين الربانية والموازين المادية الزائفة.",
    138: "الصبرُ على البلاء مع حسن الاحتساب: فالصبر ليس خمولًا ولا استسلامًا، بل عبوديةٌ عظيمة تثمر الثبات، وتمنع القلب من الانكسار، وتربط العبد بوعد الله ﷻ في الشدة والرخاء.",
    139: "الإكثارُ من الدعاء والافتقار إلى الله ﷻ: فإن القلوب لا تثبت إلا بتثبيته، ولا تنكشف الكروب إلا برحمته، ومن صدق في اللجأ إليه آواه، ومن أدام قرع بابه لم يُحرَم فضله.",
    140: "إظهارُ محاسن الإسلام في واقع الحياة: وذلك بحسن الخلق، وصدق المعاملة، والقيام بالحق، حتى يكون المسلم شاهدًا لدينه بسيرته قبل كلامه، وداعيةً إلى ربه بثباته وعدله.",
    141: "حفظُ الهوية الإيمانية في النفوس والبيوت: بتربية الأبناء على تعظيم الوحي، والاعتزاز بالإسلام، والبراءة من الهزيمة النفسية، حتى تنشأ الأجيال على الثقة بدينها والطمأنينة إلى ربها.",
    142: "فإذا قامت هذه المعاني في القلوب، واستقرت هذه الأصول في النفوس، كان المسلم أعزَّ بالله، وأثبتَ على الحق، وأبعدَ عن الهزيمة النفسية، وأقربَ إلى الفلاح في الدنيا والآخرة، بإذن الله تعالى.",
}

PARA_SPLIT_OVERRIDES = {
    42: [
        "يبرز صبر إبراهيم الخليل عليه السلام في مقام الدعوة حين واجه قومًا جمعوا له بين السخرية والتكذيب والبطش، فلم يثنه ذلك عن بيان التوحيد، ولا عن إظهار تهافت الشرك بالحجة والبرهان.",
        "فلم يكن ثباته عليه السلام ثباتَ موقفٍ عابر، بل ثباتَ قلبٍ ممتلئٍ بمعرفة الله، مستيقنٍ أنَّ الحق لا يضيع وإن تكاثر عليه أهل الباطل. ولهذا مضى في دعوته غير هيّابٍ من كثرتهم، ولا مستوحشٍ من قلّة الناصر.",
        "ثم انتهى أذى قومه إلى محاولة إحراقه، فكان توكله على ربِّه أعظم من كيدهم، فجعل الله النار عليه بردًا وسلامًا. وفي قصته أوضح شاهد على أنَّ الثبات على التوحيد، مع الصبر على الأذى، من أعظم أسباب النجاة والرفعة.",
    ],
    46: [
        "وتتجلى في سيرة شعيب عليه السلام صورةٌ أخرى من صور الثبات؛ إذ دعا قومه إلى عبادة الله، ونهاهم عن الظلم والبخس والفساد، فكان جزاؤه منهم التكذيب والاستخفاف والتهديد.",
        "وقد سجّل القرآن نظرتهم المادية القاصرة حين رأوا فيه ضعفًا، ولم يروا ما معه من سلطان الحق وقوة الحجة. ومع ذلك لم يتراجع عليه السلام، بل واصل البلاغ، وأقام الحجة، وصبر على أذى قومه صبر من يعلم أن العاقبة للمتقين.",
        "ومن هنا تظهر قيمة هذا النموذج في سياق البحث؛ فليس معيار الرفعة ما يراه الناس من أسباب القوة الظاهرة، وإنما ما يرزقه الله لعبده من ثباتٍ على الحق، وقيامٍ بوظيفة البلاغ، وصبرٍ على تبعات الطريق.",
    ],
    49: [
        "وأما يوسف عليه السلام فقصته من أظهر قصص القرآن في اجتماع المحن وتتابع الابتلاءات، مع بقاء القلب ثابتًا مطمئنًا بوعد الله. فقد انتقل من محنة الإخوة وكيدهم، إلى ظلمة الجب، ثم إلى الرق، ثم إلى السجن ظلمًا وعدوانًا.",
        "ومع هذا التتابع في الشدائد لم يفقد عليه السلام صفاء التوحيد، ولا حسن الظن بربه، ولا القيام بواجب الدعوة؛ بل ظهرت عبوديته لله في كل طور من أطوار المحنة، حتى في السجن جعله مقامَ دعوةٍ وتعليمٍ وإرشاد.",
        "ثم كانت العاقبة أن نقله الله من حال الاستضعاف إلى حال التمكين، ليبقى خبره شاهدًا على أنَّ البلاء قد يكون طريقًا إلى الرفعة، وأنَّ الصبر المقرون باليقين من أعظم أسباب الفرج وحسن العاقبة.",
    ],
    58: [
        "ويظهر صبر موسى عليه السلام في تعدد ميادين الابتلاء التي واجهها؛ فقد أوذي من فرعون وجنده، وكُذِّب وسُخِر منه، كما لقي من قومه ألوانًا من العنت والاعتراض والمراجعة.",
        "ومع هذا كله ظل قائمًا بأمر الله، ماضيًا في البلاغ، لا يصرفه عن الحق استكبارُ الجبابرة، ولا يفتّ في عضده استبطاءُ القوم أو قلة استجابتهم. وهذا من أعظم ما يربي المؤمن على أن الثبات ليس رهينًا بسهولة الطريق، بل بحسن الصلة بالله تعالى.",
        "ولهذا كان في خبر موسى عليه السلام عزاءٌ للمصلحين، وتثبيتٌ لمن يلقون من أذى الناس أو استعلائهم؛ فإن من سار على طريق الحق فلابد أن يلقى من الامتحان بقدر ما يصدق في حمله.",
    ],
    64: [
        "وفي قصة لوط عليه السلام يتجلى ثبات النبي الداعي في وجه فسادٍ مستعلن، ومنكرٍ قد تواطأ عليه المجتمع، حتى صار أهل الباطل يستنكرون الطهر، ويضيقون بأهل الاستقامة.",
        "وقد واجه لوط عليه السلام قومه بالإنكار والبيان، وصبر على استهزائهم وتهديدهم، حتى بلغ به الأمر أن تمنى لو كانت له قوةٌ تدفع أذاهم أو ركنٌ شديدٌ يأوي إليه. ومع ذلك لم يترك مقام البلاغ، ولم يساوم على الحق، ولم يلبّس في الحكم على الفاحشة.",
        "وفي هذا النموذج تربيةٌ عظيمة على أن غربة الحق لا تعني ضعفه، وأن استعلاء أهل الباطل في لحظة من اللحظات لا يسلب أهل الإيمان كرامتهم عند الله، ما داموا قائمين بحقه صابرين على أمره.",
    ],
}


def tex_escape(text: str) -> str:
    replacements = {
        "\\": r"\textbackslash{}",
        "{": r"\{",
        "}": r"\}",
        "&": r"\&",
        "%": r"\%",
        "$": r"\$",
        "#": r"\#",
        "_": r"\_",
    }
    for old, new in replacements.items():
        text = text.replace(old, new)
    return text


def clean_text(text: str) -> str:
    text = text.replace("() ", "")
    text = text.replace("()", "")
    text = text.replace("ﵟ", "﴿")
    text = text.replace("ﵞ", "﴾")
    text = text.replace("هوِيَّتِهِمْ", "هُوِيَّتِهِمْ")
    text = text.replace("هوِيَّتِهِم", "هُوِيَّتِهِم")
    text = text.replace("اللهٌ", "الله")
    text = text.replace("([FN:", "[FN:")
    text = text.replace(")]", "]")
    text = re.sub(r"\s+", " ", text).strip()
    return text


def extract_docx() -> tuple[dict[int, str], dict[int, str]]:
    paragraphs: dict[int, str] = {}
    footnotes: dict[int, str] = {}

    with zipfile.ZipFile(DOCX_PATH) as zf:
        doc = ET.fromstring(zf.read("word/document.xml"))
        foot = ET.fromstring(zf.read("word/footnotes.xml"))

        for fn in foot.findall("w:footnote", NS):
            fid = fn.attrib.get(f"{{{NS['w']}}}id")
            if fid in {"-1", "0", None}:
                continue
            pieces = []
            for p in fn.findall("w:p", NS):
                seg = []
                for node in p.iter():
                    tag = node.tag.rsplit("}", 1)[-1] if "}" in node.tag else node.tag
                    if tag == "t":
                        seg.append(node.text or "")
                    elif tag == "sym":
                        ch = node.attrib.get(f"{{{NS['w']}}}char")
                        seg.append(SYMS.get(ch, f"<SYM:{ch}>"))
                line = clean_text("".join(seg))
                if line:
                    pieces.append(line)
            footnotes[int(fid)] = clean_text(" ".join(pieces))

        for idx, p in enumerate(doc.findall(".//w:body/w:p", NS), 1):
            seg = []
            for child in p:
                tag = child.tag.rsplit("}", 1)[-1]
                if tag == "r":
                    for node in child:
                        ntag = node.tag.rsplit("}", 1)[-1]
                        if ntag == "t":
                            seg.append(node.text or "")
                        elif ntag == "tab":
                            seg.append(" ")
                        elif ntag == "sym":
                            ch = node.attrib.get(f"{{{NS['w']}}}char")
                            seg.append(SYMS.get(ch, f"<SYM:{ch}>"))
                        elif ntag == "footnoteReference":
                            fid = node.attrib.get(f"{{{NS['w']}}}id")
                            seg.append(f"[FN:{fid}]")
                        elif ntag == "lastRenderedPageBreak":
                            seg.append(" ")
            line = clean_text("".join(seg))
            if line:
                paragraphs[idx] = line

    return paragraphs, footnotes


def apply_overrides(paragraphs: dict[int, str]) -> dict[int, str | list[str]]:
    merged = dict(paragraphs)
    for idx, text in PARA_OVERRIDES.items():
        merged[idx] = clean_text(text)
    for idx, parts in PARA_SPLIT_OVERRIDES.items():
        merged[idx] = [clean_text(part) for part in parts]
    return merged


def attach_footnotes(text: str, footnotes: dict[int, str]) -> str:
    notes: dict[str, str] = {}

    def repl(match: re.Match[str]) -> str:
        fid = int(match.group(1))
        note = tex_escape(footnotes.get(fid, ""))
        marker = f"@@FN{fid}@@"
        notes[marker] = rf"\footnote{{{note}}}"
        return marker

    escaped = tex_escape(re.sub(r"\[FN:(\d+)\]", repl, text))
    for marker, footnote in notes.items():
        escaped = escaped.replace(marker, footnote)
    return escaped


def fmt_text(text: str, footnotes: dict[int, str]) -> str:
    return attach_footnotes(clean_text(text), footnotes)


def qblock(text: str, footnotes: dict[int, str]) -> str:
    return (
        "\\begin{quranblock}\n"
        + fmt_text(text, footnotes)
        + "\n\\end{quranblock}\n"
    )


def pblock(text: str | list[str], footnotes: dict[int, str]) -> str:
    if isinstance(text, list):
        return "\n\n".join(fmt_text(part, footnotes) for part in text) + "\n"
    return fmt_text(text, footnotes) + "\n"


def hadithblock(text: str, footnotes: dict[int, str]) -> str:
    return (
        "\\begin{hadithblock}\n"
        + fmt_text(text, footnotes)
        + "\n\\end{hadithblock}\n"
    )


def quoteheading(text: str, footnotes: dict[int, str]) -> str:
    return rf"\begin{{sourceintro}}{fmt_text(text, footnotes)}\end{{sourceintro}}" + "\n"


def build_body(paras: dict[int, str], footnotes: dict[int, str]) -> str:
    p = paras
    blocks: list[str] = []

    def add(text: str) -> None:
        blocks.append(text)

    add(r"\frontmatter")
    add(rf"\BookCover{{{tex_escape(p[1])}}}{{بحثٌ في مكانةِ المسلمِ عند اللهِ تعالى، ومعاني الثباتِ زمنَ الابتلاء}}")
    add(r"\tableofcontents")
    add(r"\clearpage")
    add(r"\begin{openingpage}")
    add(qblock("ﵟوَلَا تَهِنُواْ وَلَا تَحۡزَنُواْ وَأَنتُمُ ٱلۡأَعۡلَوۡنَ إِن كُنتُم مُّؤۡمِنِينَ ١٣٩ﵞ [آل عمران: 139]", footnotes))
    add(r"\end{openingpage}")
    add(r"\mainmatter")

    add(rf"\chapter{{{tex_escape(p[2])}}}")
    add(pblock(p[3], footnotes))

    add(rf"\chapter{{{tex_escape(p[4])}}}")
    add(rf"\section{{{tex_escape(p[5])}}}")
    add(pblock(p[6], footnotes))
    add(qblock(p[7], footnotes))
    add(quoteheading(p[8], footnotes))
    add(pblock(p[9], footnotes))
    for idx in range(10, 15):
        add(qblock(p[idx], footnotes))
    for idx in (15, 17, 27, 29, 31, 33, 35, 37, 39):
        add(rf"\subsection*{{{tex_escape(p[idx])}}}")
        add(hadithblock(p[idx + 1], footnotes))
    add(quoteheading(p[19], footnotes))
    for idx in range(20, 27):
        add(pblock(p[idx], footnotes))

    add(rf"\section{{{tex_escape(p[41])}}}")
    add(pblock(p[42], footnotes))
    add(qblock(p[43], footnotes))
    add(qblock(p[44], footnotes))

    add(rf"\section{{{tex_escape(p[45])}}}")
    add(pblock(p[46], footnotes))
    add(qblock(p[47], footnotes))

    add(rf"\section{{{tex_escape(p[48])}}}")
    add(pblock(p[49], footnotes))
    add(hadithblock(p[50], footnotes))
    for idx in range(51, 57):
        add(qblock(p[idx], footnotes))

    add(rf"\section{{{tex_escape(p[57])}}}")
    add(pblock(p[58], footnotes))
    for idx in range(59, 62):
        add(qblock(p[idx], footnotes))
    add(hadithblock(p[62], footnotes))

    add(rf"\section{{{tex_escape(p[63])}}}")
    add(pblock(p[64], footnotes))
    add(qblock(p[65], footnotes))
    add(qblock(p[66], footnotes))
    add(hadithblock(p[67], footnotes))

    add(rf"\chapter{{{tex_escape(p[68])}}}")
    add(pblock(p[69], footnotes))
    add(rf"\section{{{tex_escape(p[70])}}}")
    add(pblock(p[71], footnotes))
    add(hadithblock(p[72], footnotes))

    add(r"\chapter{في معاني الابتلاء والكرامة}")
    add(rf"\section{{{tex_escape(p[73])}}}")
    add(qblock(p[74], footnotes))
    add(qblock(p[75], footnotes))
    add(quoteheading(p[76], footnotes))
    for idx in range(77, 84):
        add(pblock(p[idx], footnotes))
    add(qblock(p[84], footnotes))

    add(rf"\section{{{tex_escape(p[85])}}}")
    add(pblock(p[86], footnotes))
    add(pblock(p[87], footnotes))
    add(qblock(p[88], footnotes))
    add(hadithblock(p[89], footnotes))
    add(pblock(p[90], footnotes))
    add(pblock(p[91], footnotes))
    add(qblock(p[92], footnotes))
    add(pblock(p[93], footnotes))
    add(qblock(p[94], footnotes))
    add(pblock(p[95], footnotes))
    add(qblock(p[96], footnotes))
    add(qblock(p[97], footnotes))

    add(rf"\section{{{tex_escape(p[98])}}}")
    add(pblock(p[99], footnotes))
    add(qblock(p[100], footnotes))
    add(quoteheading(p[101], footnotes))
    add(pblock(p[102], footnotes))
    add(qblock(p[103], footnotes))
    add(qblock(p[104], footnotes))
    add(qblock(p[105] + " " + p[106], footnotes))
    add(qblock(p[107], footnotes))
    add(qblock(p[108], footnotes))

    add(rf"\section{{{tex_escape(p[109])}}}")
    add(qblock(p[110], footnotes))

    add(rf"\chapter{{{tex_escape(p[111])}}}")
    add(rf"\section{{{tex_escape(p[112])}}}")
    for idx in range(113, 117):
        add(qblock(p[idx], footnotes))
    add(rf"\section{{{tex_escape(p[117])}}}")
    add(qblock(p[118], footnotes))
    add(rf"\section{{{tex_escape(p[119])}}}")
    add(qblock(p[120], footnotes))
    add(hadithblock(p[121], footnotes))
    add(quoteheading(p[122], footnotes))
    add(pblock(p[123], footnotes))

    add(rf"\chapter{{{tex_escape(p[124])}}}")
    add(qblock(p[125], footnotes))
    add(rf"\section{{{tex_escape(p[126])}}}")
    add(r"\begin{bullets}")
    for idx in range(127, 132):
        item = p[idx]
        if item.startswith("- "):
            item = item[2:]
        add(r"\item " + fmt_text(item, footnotes))
    add(r"\end{bullets}")

    add(rf"\chapter{{{tex_escape(p[132])}}}")
    for idx in range(133, 136):
        add(pblock(p[idx], footnotes))
    add(r"\begin{bullets}")
    for idx in range(136, 142):
        item = p[idx]
        if item.startswith("- "):
            item = item[2:]
        add(r"\item " + fmt_text(item, footnotes))
    add(r"\end{bullets}")
    add(pblock(p[142], footnotes))
    add(r"\begin{closingline}")
    add(fmt_text(p[143], footnotes))
    add(r"\end{closingline}")

    return "\n".join(blocks)


def build_document(body: str) -> str:
    return rf"""\documentclass[12pt,openany]{{book}}
\usepackage[a5paper,margin=16mm,headsep=8mm,footskip=10mm]{{geometry}}
\usepackage{{fontspec}}
\usepackage{{polyglossia}}
\usepackage{{graphicx}}
\usepackage{{xcolor}}
\usepackage{{tikz}}
\usepackage{{fancyhdr}}
\usepackage{{setspace}}
\usepackage{{enumitem}}
\usepackage{{microtype}}
\usepackage{{hyperref}}
\hypersetup{{hidelinks}}

\setdefaultlanguage{{arabic}}
\setotherlanguage{{english}}

\newfontfamily\arabicfont[Script=Arabic,Scale=1.08]{{Amiri}}
\newfontfamily\arabicfontsf[Script=Arabic,Scale=1.0]{{Noto Sans Arabic}}
\newfontfamily\quranfont[Script=Arabic,Scale=1.18]{{Amiri}}
\newfontfamily\titlefont[Script=Arabic,Scale=1.15]{{Amiri}}

\definecolor{{paper}}{{HTML}}{{F8F1E3}}
\definecolor{{ink}}{{HTML}}{{2A2118}}
\definecolor{{gold}}{{HTML}}{{8C6A2A}}
\definecolor{{line}}{{HTML}}{{C9B07D}}
\definecolor{{soft}}{{HTML}}{{EFE4CC}}

\setstretch{{1.28}}
\setlength{{\headheight}}{{22pt}}
\setlength{{\parindent}}{{1.2em}}
\setlength{{\parskip}}{{0.5em}}
\setlist[itemize]{{label=--,itemsep=0.6em,leftmargin=1.6em}}

\pagestyle{{fancy}}
\fancyhf{{}}
\fancyhead[LE,RO]{{\small\thepage}}
\fancyhead[RE]{{\small المسلم لا يهون على الله وإن هان على الناس}}
\fancyhead[LO]{{\small \nouppercase{{\leftmark}}}}
\renewcommand{{\headrulewidth}}{{0.2pt}}
\renewcommand{{\footrulewidth}}{{0pt}}

\newenvironment{{quranblock}}
  {{\par\medskip\begin{{center}}\begin{{minipage}}{{0.92\linewidth}}\centering\quranfont\large\color{{ink}}\hrule\vspace{{0.65em}}}}
  {{\vspace{{0.65em}}\hrule\end{{minipage}}\end{{center}}\medskip}}

\newenvironment{{hadithblock}}
  {{\par\medskip\noindent\begin{{minipage}}{{\linewidth}}\hrule height 0.7pt \vspace{{0.6em}}\small}}
  {{\vspace{{0.4em}}\hrule height 0.7pt\end{{minipage}}\par\medskip}}

\newenvironment{{sourceintro}}
  {{\par\medskip\noindent\titlefont\color{{gold}}}}
  {{\par\smallskip}}

\newenvironment{{bullets}}
  {{\begin{{itemize}}}}
  {{\end{{itemize}}}}

\newenvironment{{openingpage}}
  {{\thispagestyle{{empty}}\vspace*{{0.22\textheight}}}}
  {{\clearpage}}

\newenvironment{{closingline}}
  {{\begin{{center}}\titlefont}}
  {{\end{{center}}}}

\newcommand{{\BookCover}}[2]{{
  \begin{{titlepage}}
  \thispagestyle{{empty}}
  \begin{{tikzpicture}}[remember picture,overlay]
    \node at (current page.center) {{\includegraphics[width=\paperwidth,height=\paperheight]{{assets/ppt-cover-bg.jpg}}}};
    \fill[paper,opacity=0.86] (current page.south west) rectangle (current page.north east);
    \draw[line width=1pt,color=line] ([xshift=10mm,yshift=10mm]current page.south west) rectangle ([xshift=-10mm,yshift=-10mm]current page.north east);
  \end{{tikzpicture}}
  \vspace*{{0.16\textheight}}
  \begin{{center}}
    {{\titlefont\fontsize{{26}}{{34}}\selectfont\color{{ink}} #1\par}}
    \vspace{{1.2em}}
    {{\Large\color{{gold}} #2\par}}
    \vfill
    {{\quranfont\large\color{{ink}} ﴿وَلَا تَهِنُواْ وَلَا تَحۡزَنُواْ وَأَنتُمُ ٱلۡأَعۡلَوۡنَ إِن كُنتُم مُّؤۡمِنِينَ﴾\par}}
    \vspace{{0.6em}}
    {{\small\color{{ink}} [آل عمران: 139]\par}}
  \end{{center}}
  \end{{titlepage}}
}}

\begin{{document}}
{body}
\end{{document}}
"""


def main() -> None:
    paras, footnotes = extract_docx()
    paras = apply_overrides(paras)
    body = build_body(paras, footnotes)
    tex = build_document(body)
    OUTPUT_PATH.write_text(tex, encoding="utf-8")
    print(OUTPUT_PATH.name)


if __name__ == "__main__":
    main()
