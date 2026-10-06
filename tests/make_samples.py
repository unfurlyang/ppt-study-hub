# -*- coding: utf-8 -*-
"""生成多格式测试样本：DOCX / PDF / TXT / MD
DOCX 与 TXT 用中文物理内容，PDF 用英文（fpdf2 核心字体不含中文）。
"""
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "samples")
os.makedirs(OUT, exist_ok=True)


def make_docx():
    from docx import Document
    doc = Document()
    doc.add_heading("牛顿第一定律", level=0)
    paras = [
        "本讲义介绍牛顿第一定律的内容、物理意义及其在生活中的应用。同学们需要重点掌握惯性的概念，并能够运用惯性解释常见的物理现象。",
        "一、定律内容：一切物体总保持匀速直线运动状态或静止状态，除非作用在它上面的力迫使它改变这种状态。这就是著名的牛顿第一定律，也称为惯性定律。",
        "二、伽利略的理想斜面实验。伽利略通过理想斜面实验推理指出：如果摩擦力可以完全忽略，小球将上升到原来的高度；当对接斜面变为水平时，小球将永远运动下去。",
        "这个实验的意义在于，它开创了以实验事实为基础、通过逻辑推理得出物理规律的科学方法，是物理学史上里程碑式的工作。",
        "三、惯性的概念。物体保持原有运动状态不变的性质叫做惯性。一切物体都具有惯性，惯性是物体的固有属性，质量是惯性大小的唯一量度。",
        "四、生活中的惯性现象举例：",
        "1. 汽车急刹车时，乘客身体向前倾，这是因为乘客下半身随车减速，而上半身由于惯性仍保持原来的运动状态。",
        "2. 锤头松了，把锤柄在凳子上撞击几下，锤头就紧套在锤柄上，这是利用了锤头的惯性。",
        "3. 跳远运动员起跳前助跑，是利用身体的惯性来提高成绩。",
        "五、需要注意的误区：力不是维持物体运动的原因，而是改变物体运动状态的原因。亚里士多德的错误观点统治了近两千年，直到伽利略和牛顿才被纠正。",
        "六、牛顿第一定律与牛顿第二定律的关系：第一定律定性地给出了力与运动的关系，并定义了惯性；第二定律则定量地给出 F=ma。第一定律是第二定律在合力为零时的特殊情形，但它的价值远不止于此。",
        "课堂小结：牛顿第一定律揭示了力与运动的正确关系；惯性是物体的固有属性，质量越大惯性越大；解释惯性现象要抓住「原来状态」与「受力改变」两个要点。",
    ]
    for p in paras:
        doc.add_paragraph(p)
    path = os.path.join(OUT, "sample.docx")
    doc.save(path)
    return path


def make_pdf():
    from fpdf import FPDF
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    pdf.set_font("Helvetica", "B", 18)
    pdf.cell(0, 12, "Newton's First Law of Motion", new_x="LMARGIN", new_y="NEXT", align="C")
    pdf.ln(4)
    pdf.set_font("Helvetica", "", 12)
    body = [
        "This lecture note summarizes Newton's First Law and its applications in daily life.",
        "1. Statement of the Law: An object at rest stays at rest, and an object in motion stays in motion with the same speed and in the same direction, unless acted upon by an unbalanced external force. This is also called the Law of Inertia.",
        "2. Galileo's Thought Experiment: Galileo rolled a ball down one incline and observed that it climbed the opposite incline to nearly the same height. He reasoned that if the second incline were made horizontal, the ball would keep rolling forever, assuming no friction.",
        "This experiment established a new scientific method: deriving physical laws from experimental facts combined with logical reasoning.",
        "3. The Concept of Inertia: Inertia is the tendency of an object to resist changes in its state of motion. All objects have inertia. Mass is the only measure of inertia. The larger the mass, the harder it is to change the motion of the object.",
        "4. Everyday Examples: When a car brakes suddenly, passengers lurch forward because their bodies tend to keep moving. A hammer head can be tightened by striking the handle against a bench, because the heavy head keeps moving when the handle stops.",
        "5. Common Misconception: Force is NOT needed to maintain motion. Force is needed to CHANGE motion. Aristotle's view was corrected by Galileo and Newton after nearly two thousand years.",
        "6. Key Takeaways: The First Law defines inertial frames and reveals the correct relation between force and motion. It is the qualitative foundation on which the Second Law F = ma is built.",
        "Review Questions: State the First Law in your own words. Explain why passengers lean forward when a bus stops suddenly. Why is mass, not volume or weight, the measure of inertia?",
    ]
    for para in body:
        pdf.multi_cell(0, 7, para)
        pdf.ln(1)
    path = os.path.join(OUT, "sample.pdf")
    pdf.output(path)
    return path


def make_txt():
    text = """牛顿第二定律课堂讲义

一、定律内容：物体的加速度跟所受的合外力成正比，跟物体的质量成反比，加速度的方向跟合外力的方向相同。公式表达为 F = ma。

二、公式的理解要点。F 指的是物体所受的合外力，而不是某一个力；m 是物体的质量，反映物体惯性的大小；a 与 F 是瞬时对应关系，合力变化，加速度随之变化。

三、单位的规定。在国际单位制中，力的单位是牛顿，符号 N。使质量为 1 千克的物体产生 1 米每二次方秒加速度的力，定义为 1 牛顿。

四、应用解题的一般步骤：第一，确定研究对象；第二，进行受力分析并画出受力图；第三，建立直角坐标系并正交分解；第四，列出 F=ma 的分量式求解。

五、典型例题：质量为 2 千克的物体在水平面上受到 10 牛顿的水平拉力，物体与地面间的滑动摩擦力为 4 牛顿，求物体的加速度。解：合力 F = 10 - 4 = 6 牛顿，加速度 a = F/m = 6/2 = 3 米每二次方秒，方向与拉力相同。

课堂小结：牛顿第二定律是动力学的核心规律，解题时关键是正确求出合外力，并注意各量的方向关系。
"""
    path = os.path.join(OUT, "sample.txt")
    with open(path, "w", encoding="utf-8") as f:
        f.write(text)
    return path


def make_md():
    text = """# 匀变速直线运动

## 一、基本概念

沿着一条直线且加速度不变的运动，叫做匀变速直线运动。速度随时间均匀增加的为匀加速，均匀减小的为匀减速。

## 二、三个基本公式

- 速度公式：v = v0 + at
- 位移公式：x = v0*t + (1/2)*a*t^2
- 速度位移关系：v^2 - v0^2 = 2ax

## 三、重要推论

匀变速直线运动中，任意相邻相等时间间隔内的位移差为恒量，即 x2 - x1 = a*T^2。这是打点计时器纸带处理的核心依据。

## 四、图像分析

v-t 图像中，图线的斜率表示加速度，图线与时间轴围成的面积表示位移。斜率为正表示加速，为负表示减速。

## 五、课堂小结

掌握三个基本公式及其适用条件，学会利用 v-t 图像分析运动过程，是本章的核心要求。解题时注意先规定正方向。
"""
    path = os.path.join(OUT, "sample.md")
    with open(path, "w", encoding="utf-8") as f:
        f.write(text)
    return path


def verify(paths):
    from docx import Document
    doc = Document(paths["docx"])
    n_para = len([p for p in doc.paragraphs if p.text.strip()])
    print("[verify] DOCX 非空段落数 =", n_para)
    from pypdf import PdfReader
    r = PdfReader(paths["pdf"])
    txt = "".join((pg.extract_text() or "") for pg in r.pages)
    print("[verify] PDF 页数 =", len(r.pages), "| 提取字符数 =", len(txt))
    with open(paths["txt"], encoding="utf-8") as f:
        t = f.read()
    print("[verify] TXT 字符数 =", len(t), "| 空行分段数 =",
          len([s for s in t.split("\n\n") if s.strip()]))
    with open(paths["md"], encoding="utf-8") as f:
        m = f.read()
    print("[verify] MD 字符数 =", len(m))
    for k, p in paths.items():
        print(f"[file] {k}: {p} ({os.path.getsize(p)} bytes)")


if __name__ == "__main__":
    paths = {
        "docx": make_docx(),
        "pdf": make_pdf(),
        "txt": make_txt(),
        "md": make_md(),
    }
    verify(paths)
    print("ALL SAMPLES GENERATED")
