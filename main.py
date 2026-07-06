"""
B站排行榜爬虫 - Kivy Android App (简化版)
"""
from kivy.app import App
from kivy.uix.label import Label


class TestApp(App):
    def build(self):
        return Label(text="Hello Bilibili!")


if __name__ == "__main__":
    TestApp().run()
