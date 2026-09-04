# -*- coding: utf-8 -*-
"""GUI 启动入口（附加测试按钮补丁）"""
import main
import gui_patch
gui_patch.apply()

app = main.tb.Window(themename='flatly')
main.DiffViewer(app)
app.mainloop()
