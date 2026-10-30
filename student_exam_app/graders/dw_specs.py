"""DW-16/213 题组 5 个小题的评分规范（依据 结果素材 与 考试素材 的差异整理）。

每个小题 6 个检查项合计 100 分。文件名不区分大小写。
"""

SPECS = {
    '213-1': {
        'title': '网页操作题 213-1（关爱珍惜动物）',
        'items': [
            {'name': '框架结构', 'points': 20, 'checks': [
                ('attr', 'index.html', 'frameset', 'rows', '12%,*'),
                ('frame', 'index.html', 'middle.html'),
                ('frame', 'index.html', 'right.html'),
            ], 'labels': [
                '上框架高度改为 12%',
                '中部框架指向 middle.html',
                '新增右框架指向 right.html',
            ]},
            {'name': '上页标题与文字', 'points': 20, 'checks': [
                ('title', 'top.html', '关爱'),
                ('text', 'top.html', '关爱珍惜动物'),
                ('css', 'top.html', '幼圆', '36px', '#ff9900'),
                ('css', 'top.html', 'background-color:#009900'),
            ], 'labels': [
                '标题改为「关爱」',
                '插入文字「关爱珍惜动物」',
                'CSS 幼圆 36px #ff9900',
                '网页背景色 #009900',
            ]},
            {'name': '左页背景', 'points': 15, 'checks': [
                ('body', 'left.html', 'bgcolor', '#ff9900'),
                ('body', 'left.html', 'link', '#ff0000'),
            ], 'labels': [
                '左页背景 #FF9900',
                '链接颜色 #FF0000',
            ]},
            {'name': '两栖类链接', 'points': 20, 'checks': [
                ('tag', 'left.html', 'a', {'href': 'middle.html#两栖'}),
            ], 'labels': [
                '「两栖类」链接到 middle.html#两栖',
            ]},
            {'name': '书签名锚点', 'points': 15, 'checks': [
                ('anchor', 'middle.html', '两栖'),
            ], 'labels': [
                'middle.html 设置书签「两栖」',
            ]},
            {'name': '水平线与按钮', 'points': 10, 'checks': [
                ('tag', 'left.html', 'hr', {'size': '1', 'noshade': True}),
                ('tag', 'left.html', 'input', {'type': 'submit', 'value': '更多'}),
            ], 'labels': [
                '插入水平线（高1，无阴影）',
                '插入「更多」按钮',
            ]},
        ],
    },
    '213-2': {
        'title': '网页操作题 213-2（唐诗欣赏）',
        'items': [
            {'name': '顶部框架', 'points': 20, 'checks': [
                ('attr', 'index.html', 'frameset', 'rows', '100,*'),
                ('frame', 'index.html', 'top1.html'),
                ('frame', 'index.html', 'top2.html'),
            ], 'labels': [
                '上框架高度改为 100 像素',
                '上部嵌套框架含 TOP1.html',
                '上部嵌套框架含 TOP2.html',
            ]},
            {'name': '上部列宽', 'points': 15, 'checks': [
                ('attr', 'index.html', 'frameset', 'cols', '512,512'),
            ], 'labels': [
                '上部框架列宽 512、512',
            ]},
            {'name': '下部框架', 'points': 20, 'checks': [
                ('attr', 'index.html', 'frameset', 'cols', '200,*'),
                ('frame', 'index.html', 'left.html'),
            ], 'labels': [
                '下部框架列宽 200、*',
                '左框架指向 left.html',
            ]},
            {'name': '诗文链接', 'points': 15, 'checks': [
                ('tag', 'right.html', 'a', {'href': '#'}),
            ], 'labels': [
                '「人生情最重」设空链接',
            ]},
            {'name': '书签', 'points': 15, 'checks': [
                ('anchor', 'right.html', 'youqing'),
            ], 'labels': [
                '「友情是帆」前设书签 youqing',
            ]},
            {'name': '新建左页', 'points': 15, 'checks': [
                ('file', 'left.html'),
            ], 'labels': [
                '新建 left.html 页面',
            ]},
        ],
    },
    '213-3': {
        'title': '网页操作题 213-3（健康与读书）',
        'items': [
            {'name': '框架结构', 'points': 20, 'checks': [
                ('attr', 'index.html', 'frameset', 'rows', '170,*'),
                ('attr', 'index.html', 'frameset', 'rows', '70%,30%'),
                ('frame', 'index.html', 'bottom.html'),
            ], 'labels': [
                '上框架高度 170 像素',
                '中部嵌套框架 70%、30%',
                '下框架指向 bottom.html',
            ]},
            {'name': '左页标题与背景', 'points': 15, 'checks': [
                ('title', 'left.html', '目录'),
                ('body', 'left.html', 'bgcolor', '#ffdece'),
            ], 'labels': [
                '标题改为「目录」',
                '背景色 #FFDECE',
            ]},
            {'name': '水平线与提示', 'points': 15, 'checks': [
                ('tag', 'left.html', 'hr', {'size': '1', 'noshade': True}),
                ('text', 'left.html', '您认为在哪些方面自己需要加强'),
            ], 'labels': [
                '顶部插入水平线（高1，无阴影）',
                '输入提示文字',
            ]},
            {'name': '下拉菜单', 'points': 25, 'checks': [
                ('tag', 'left.html', 'select', {'size': '1'}),
                ('select', 'left.html', ['运动', '饮食', '读书', '其他'], '读书'),
            ], 'labels': [
                '插入下拉菜单',
                '选项 运动/饮食/读书/其他，默认选中「读书」',
            ]},
            {'name': '标题样式', 'points': 15, 'checks': [
                ('css', 'middle.html', 'font-family:"幼圆"'),
                ('css', 'middle.html', 'color:#b54e09'),
            ], 'labels': [
                '标题字体改为幼圆',
                '标题颜色 #B54E09',
            ]},
            {'name': '表格对齐', 'points': 10, 'checks': [
                ('tag', 'middle.html', 'table', {'align': 'center'}),
            ], 'labels': [
                '表格设置为居中对齐',
            ]},
        ],
    },
    '213-4': {
        'title': '网页操作题 213-4（植物世界）',
        'items': [
            {'name': '框架结构', 'points': 20, 'checks': [
                ('attr', 'index.html', 'frameset', 'rows', '130,*'),
                ('attr', 'index.html', 'frameset', 'rows', '60%,40%'),
                ('frame', 'index.html', 'bottom.html'),
            ], 'labels': [
                '上框架高度 130 像素',
                '中部嵌套框架 60%、40%',
                '下框架指向 bottom.html',
            ]},
            {'name': '页面标题', 'points': 15, 'checks': [
                ('title', 'middle.html', '植物世界'),
            ], 'labels': [
                '网页标题改为「植物世界」',
            ]},
            {'name': '表格边框', 'points': 15, 'checks': [
                ('tag', 'middle.html', 'table', {'border': '0'}),
            ], 'labels': [
                '表格边框改为 0',
            ]},
            {'name': '单元格插图', 'points': 20, 'checks': [
                ('tag', 'middle.html', 'img', {'src': 'image/word.gif', 'width': '175', 'height': '28'}),
            ], 'labels': [
                '单元格插入 image/WORD.GIF（175×28）',
            ]},
            {'name': '右对齐水平线', 'points': 15, 'checks': [
                ('tag', 'middle.html', 'hr', {'align': 'right', 'width': '100%', 'size': '3', 'noshade': True}),
            ], 'labels': [
                '插入右对齐水平线（宽100%、高3、无阴影）',
            ]},
            {'name': '上页文字与样式', 'points': 15, 'checks': [
                ('text', 'top.html', '植物世界'),
                ('css', 'top.html', '幼圆', '72px', '#ff9900'),
            ], 'labels': [
                '上页插入文字「植物世界」',
                'CSS 幼圆 72px #ff9900',
            ]},
        ],
    },
    '213-5': {
        'title': '网页操作题 213-5（中国教育网）',
        'items': [
            {'name': '框架结构', 'points': 20, 'checks': [
                ('attr', 'index.html', 'frameset', 'rows', '110,*'),
                ('attr', 'index.html', 'frameset', 'rows', '70%,30%'),
                ('frame', 'index.html', 'bottom.html'),
            ], 'labels': [
                '上框架高度 110 像素',
                '中部嵌套框架 70%、30%',
                '下框架指向 bottom.html',
            ]},
            {'name': '左页背景', 'points': 15, 'checks': [
                ('body', 'left.html', 'bgcolor', '#0099cc'),
                ('tag', 'left.html', 'hr', {'size': '1', 'noshade': True}),
            ], 'labels': [
                '背景色 #0099CC',
                '顶部插入水平线',
            ]},
            {'name': '表格边框', 'points': 15, 'checks': [
                ('tag', 'middle.html', 'table', {'border': '0'}),
            ], 'labels': [
                '表格边框改为 0',
            ]},
            {'name': '下拉菜单', 'points': 20, 'checks': [
                ('select', 'middle.html', ['小学教育', '高中教育', '大学教育', '学前教育'], '学前教育'),
            ], 'labels': [
                '插入下拉菜单，默认选中「学前教育」',
            ]},
            {'name': '标题与样式', 'points': 20, 'checks': [
                ('title', 'top.html', '中国教育网'),
                ('css', 'top.html', '幼圆', '60px', '#ff9900'),
                ('text', 'top.html', '中国教育网'),
            ], 'labels': [
                '标题改为「中国教育网」',
                'CSS 幼圆 60px #ff9900',
                '页面插入文字「中国教育网」',
            ]},
            {'name': '上页插图', 'points': 10, 'checks': [
                ('tag', 'top.html', 'img', {'src': 'image/1.gif', 'width': '465', 'height': '61'}),
            ], 'labels': [
                '插入 image/1.GIF（465×61）',
            ]},
        ],
    },
}

SPECS.update({
    '34': {
        'title': '网页操作题 034（奥运剪纸 · 60分）',
        'items': [
            {'name': '网页标题与关键字', 'points': 15, 'checks': [
                ('title', 'index.html', '奥运剪纸'),
                ('meta', 'index.html', 'keywords', '剪纸'),
            ], 'labels': ['标题为「奥运剪纸」', '关键字内容为「剪纸」']},
            {'name': '横幅图片', 'points': 20, 'checks': [
                ('tag', 'index.html', 'img', {'src': 'image/banner.jpg'}),
                ('tag', 'index.html', 'img', {'src': 'image/banner.jpg', 'alt': '中国剪纸网'}),
                ('tag', 'index.html', 'img', {'src': 'image/banner.jpg', 'hspace': '105', 'border': '1'}),
            ], 'labels': ['插入 banner.jpg', '替换文本「中国剪纸网」', '水平边距105、边框1']},
            {'name': '三栏表格', 'points': 20, 'checks': [
                ('tag', 'index.html', 'table', {'width': '773', 'border': '1', 'align': 'center',
                                                'cellpadding': '0', 'cellspacing': '0'}),
                ('text', 'index.html', '剪纸产品展示'),
                ('text', 'index.html', '剪纸产品应用'),
            ], 'labels': ['1行3列表格 宽773/边框1/居中/边距0', '第1列素材文字', '第3列素材文字']},
            {'name': '中栏图片与文件域', 'points': 15, 'checks': [
                ('tag', 'index.html', 'img', {'src': 'image/mid.jpg'}),
                ('tag', 'index.html', 'input', {'type': 'file', 'size': '16'}),
            ], 'labels': ['第2列插入 mid.jpg', '文件域字符宽度16']},
            {'name': '五图表格', 'points': 15, 'checks': [
                ('tag', 'index.html', 'table', {'width': '774', 'border': '1', 'align': 'center'}),
                ('tag', 'index.html', 'img', {'src': 'image/jianzhi01.jpg'}),
                ('tag', 'index.html', 'img', {'src': 'image/jianzhi05.jpg'}),
            ], 'labels': ['1行5列表格 宽774/居中', '插入 jianzhi01.jpg', '插入 jianzhi05.jpg']},
            {'name': '替换文本与热点', 'points': 15, 'checks': [
                ('tag', 'index.html', 'img', {'src': 'image/jianzhi01.jpg', 'alt': '中国民间剪纸'}),
                ('tag', 'index.html', 'img', {'src': 'image/jianzhi03.jpg', 'alt': '梅花剪纸'}),
                ('area', 'index.html', {'shape': 'circle', 'href': 'www.21sgy.com'}),
            ], 'labels': ['01替换文本「中国民间剪纸」', '03替换文本「梅花剪纸」', '椭圆形热点链接 21sgy.com']},
        ],
    },
    '7': {
        'title': '网页操作题 007（走进香格里拉 · 60分）',
        'items': [
            {'name': '网页标题', 'points': 10, 'checks': [
                ('title', 'index.html', '走进香格里拉'),
            ], 'labels': ['标题为「走进香格里拉」']},
            {'name': '拆分新增行图片', 'points': 15, 'checks': [
                ('tag', 'index.html', 'img', {'src': 'image/1.jpg'}),
                ('tag', 'index.html', 'img', {'src': 'image/2.jpg'}),
                ('tag', 'index.html', 'img', {'src': 'image/3.jpg'}),
            ], 'labels': ['第1行插入 1.jpg', '第2行插入 2.jpg', '第2行插入 3.jpg']},
            {'name': '两个矩形热点', 'points': 20, 'checks': [
                ('area', 'index.html', {'shape': 'rect', 'href': 'link.html'}),
                ('area', 'index.html', {'shape': 'rect', 'href': 'www.zzxs.com'}),
                ('area', 'index.html', {'alt': '那不仅是一种景观，也是一种意境'}),
            ], 'labels': ['热点链接 link.html', '热点链接 www.zzxs.com', '热点替换文本']},
            {'name': '水平线', 'points': 15, 'checks': [
                ('tag', 'index.html', 'hr', {'width': '400', 'noshade': True}),
            ], 'labels': ['插入水平线 宽400 无阴影']},
            {'name': '标题文字格式', 'points': 20, 'checks': [
                ('text', 'index.html', '悬念之旅'),
                ('tag', 'index.html', 'font', {'size': '5', 'color': '#cc33ff', 'face': '华文仿宋'}),
            ], 'labels': ['插入「悬念之旅」文字', '字体 华文仿宋/5号/#CC33FF']},
            {'name': '层与日期', 'points': 20, 'checks': [
                ('tag', 'index.html', 'div', {'id': 'layer1'}),
                ('tag', 'index.html', 'img', {'src': 'image/dqm.jpg'}),
                ('date', 'index.html'),
            ], 'labels': ['插入层 Layer1', '层内插入 dqm.jpg', '层内插入自动更新日期']},
        ],
    },
    '6': {
        'title': '网页操作题 006（中礼鲜花速递 · 60分）',
        'items': [
            {'name': '网页标题', 'points': 8, 'checks': [
                ('title', 'index.html', '中礼鲜花速递'),
            ], 'labels': ['标题为「中礼鲜花速递」']},
            {'name': '顶部图片与层', 'points': 12, 'checks': [
                ('tag', 'index.html', 'img', {'src': 'image/logo.jpg'}),
                ('tag', 'index.html', 'img', {'src': 'image/top1.gif'}),
                ('tag', 'index.html', 'div', {'id': 'layer1'}),
            ], 'labels': ['插入 logo.jpg', '插入 top1.gif', '用层放置 top1.gif']},
            {'name': '两个下拉菜单', 'points': 20, 'checks': [
                ('select', 'index.html', ['所有', '爱情', '友情', '祝福'], ''),
                ('select', 'index.html', ['80元以下', '80-200元', '200以上'], ''),
            ], 'labels': ['用途列表 所有/爱情/友情/祝福', '价格列表 80元以下/80-200元/200以上']},
            {'name': '文本框与按钮', 'points': 15, 'checks': [
                ('tag', 'index.html', 'input', {'type': 'text'}),
                ('tag', 'index.html', 'input', {'value': '搜索'}),
            ], 'labels': ['插入文本域', '按钮值「搜索」']},
            {'name': '表格与图片', 'points': 25, 'checks': [
                ('tag', 'index.html', 'table', {'border': '1', 'bordercolor': '#99cc66'}),
                ('tag', 'index.html', 'td', {'background': 'image/line9.jpg'}),
                ('tag', 'index.html', 'img', {'src': 'image/1.jpg', 'alt': '一心一意'}),
                ('tag', 'index.html', 'img', {'src': 'image/3.jpg', 'alt': '亲亲宝贝'}),
            ], 'labels': ['表格边框1/颜色#99CC66', '首行单元格背景 line9.jpg', '图片1替换文本「一心一意」',
                         '图片3替换文本「亲亲宝贝」']},
            {'name': '层内欢迎语与邮件链接', 'points': 20, 'checks': [
                ('tag', 'index.html', 'div', {'id': 'layer2'}),
                ('text', 'index.html', '中礼欢迎您的光临'),
                ('css', 'index.html', 'font-size:36px', 'color:#ff0000'),
                ('tag', 'index.html', 'a', {'href': 'mailto:jinjia@jinjiaflower.com'}),
            ], 'labels': ['插入层 Layer2', '文字「中礼欢迎您的光临」', '36px 红色', '邮件链接']},
        ],
    },
    '8': {
        'title': '网页操作题 008（北京奥运会火炬 · 60分）',
        'items': [
            {'name': '网页标题', 'points': 10, 'checks': [
                ('title', 'index.html', '北京奥运会火炬'),
            ], 'labels': ['标题为「北京奥运会火炬」']},
            {'name': '介绍文字格式', 'points': 20, 'checks': [
                ('css', 'index.html', '华文仿宋', 'font-size:36px', 'color:#6699ff'),
            ], 'labels': ['文字设为 华文仿宋/36px/#6699FF']},
            {'name': '命名锚记与链接', 'points': 20, 'checks': [
                ('anchor', 'index.html', '01'),
                ('tag', 'index.html', 'a', {'href': '#01'}),
            ], 'labels': ['锚记名称 01', '文字链接到 #01']},
            {'name': '火炬图片', 'points': 20, 'checks': [
                ('tag', 'index.html', 'img', {'src': 'image/1.jpg', 'alt': '奥运火炬'}),
                ('a_img', 'index.html', 'link.html', 'image/1.jpg'),
            ], 'labels': ['插入 1.jpg 替换文本「奥运火炬」', '图片链接 link.html']},
            {'name': '组图链接', 'points': 15, 'checks': [
                ('a_text', 'index.html', '组图', 'link.html'),
            ], 'labels': ['「组图」文字链接 link.html']},
            {'name': '自动更新日期', 'points': 15, 'checks': [
                ('date', 'index.html'),
            ], 'labels': ['插入储存时自动更新日期']},
        ],
    },
    '9': {
        'title': '网页操作题 009（瑜伽 · 60分）',
        'items': [
            {'name': '位置1 文字', 'points': 18, 'checks': [
                ('text', 'index.html', '瑜伽是一种非常古老的能量知识修炼方法'),
                ('css', 'index.html', 'font-size:24px', 'color:#ffba00'),
            ], 'labels': ['文字内容', '24px #FFBA00']},
            {'name': '位置2 文字', 'points': 14, 'checks': [
                ('text', 'index.html', '瑜伽能够改善人们生理、心理、情感和精神'),
            ], 'labels': ['文字内容']},
            {'name': '位置3 素材文字', 'points': 20, 'checks': [
                ('text', 'index.html', '瑜伽（拼音：yú'),
                ('css', 'index.html', 'font-size:24px', 'color:#ff6600'),
            ], 'labels': ['插入素材第一段', '首行 24px #FF6600']},
            {'name': '位置4 层', 'points': 18, 'checks': [
                ('tag', 'index.html', 'div', {'id': 'layer1'}),
                ('text', 'index.html', '瑜伽的分类'),
            ], 'labels': ['插入层 Layer1', '层内文字「瑜伽的分类」']},
            {'name': '位置5 五张图', 'points': 15, 'checks': [
                ('tag', 'index.html', 'img', {'src': 'image/1.jpg'}),
                ('tag', 'index.html', 'img', {'src': 'image/3.jpg'}),
                ('tag', 'index.html', 'img', {'src': 'image/5.jpg'}),
            ], 'labels': ['插入 1.jpg', '插入 3.jpg', '插入 5.jpg']},
            {'name': '新窗口链接', 'points': 15, 'checks': [
                ('tag', 'index.html', 'a', {'href': 'wy.html', 'target': '_blank'}),
            ], 'labels': ['1.jpg 链接 wy.html 且新窗口打开']},
        ],
    },
})
