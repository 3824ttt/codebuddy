import urllib.request, json

tests = [
    ("三体", "刘慈欣", "科幻;外星文明", "文革期间军方探寻外星文明的绝密计划"),
    ("活着", "余华", "乡土;人生", "地主少爷福贵嗜赌成性败光家业"),
    ("Python编程从入门到实践", "Eric Matthes", "Python;编程", "通俗易懂的Python编程入门教程"),
    ("高等数学", "同济大学", "数学;高等数学", "高数教材"),
    ("论语译注", "杨伯峻", "论语;儒家", "论语注释翻译"),
    ("时间简史", "霍金", "宇宙;物理学", "从大爆炸到黑洞"),
    ("白夜行", "东野圭吾", "推理;悬疑", "两个孩子的命运交织"),
    ("斯坦福机器学习", "Andrew Ng", "机器学习;AI", "经典ML入门教材"),
    ("本草纲目", "李时珍", "中药;中医", "古代药物学巨著"),
    ("国家为什么会失败", "阿西莫格鲁", "经济学;政治", "制度如何影响国家兴衰"),
    ("国富论", "亚当斯密", "经济学;经典", "现代经济学开山之作"),
    ("百年孤独", "马尔克斯", "魔幻现实主义;文学", "布恩迪亚家族七代人的传奇"),
    ("C++ Primer", "Stanley Lippman", "编程;C++", "C++经典教程"),
    ("人类简史", "赫拉利", "人类学;历史", "从动物到上帝"),
    ("精神分析引论", "弗洛伊德", "心理学;精神分析", "弗洛伊德代表作"),
    ("电路基础", "", "电子;电路", "本科电子技术基础教材"),
    ("苏东坡传", "林语堂", "传记;文学", "苏东坡的传奇一生"),
    ("自动控制原理", "胡寿松", "自动化;控制", "经典控制理论教材"),
    ("信号与系统", "奥本海姆", "信号;通信", "电子工程经典教材"),
    ("梦的解析", "弗洛伊德", "心理学;梦境", "精神分析学派奠基之作"),
]

print("=" * 70)
print("智能分类测试结果")
print("=" * 70)

for title, author, keywords, desc in tests:
    data = json.dumps({
        "title": title,
        "author": author,
        "keywords": keywords,
        "description": desc,
    }).encode()
    req = urllib.request.Request(
        'http://localhost:5000/api/admin/classify/batch?books=1',
        data=json.dumps({}).encode(),
        headers={'Content-Type':'application/json', 'X-Admin-Token':'admin123'},
        method='POST'
    )
    
    # Use direct classify via test endpoint
    from classifier import classify_book
    result = classify_book(title=title, author=author, keywords=keywords, description=desc)
    
    status = "✓" if result['confidence'] >= 0.3 else "△"
    print(f"  {status} 《{title}》 → {result['category']} | {result.get('sub_category','')} | 中图:{result['clc_number']} | 置信度:{result['confidence']:.2f}")

print("\n所有分类均通过规则引擎完成（离线本地运算）")
