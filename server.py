from http.server import HTTPServer, BaseHTTPRequestHandler
from urllib.parse import unquote
import json
import subprocess
import os
import traceback

class PresentationHandler(BaseHTTPRequestHandler):
    def log_message(self, format, *args):
        print(f"[SERVER] {args[0]}")

    def send_json_response(self, status_code, data):
        """Вспомогательный метод для гарантированной отправки JSON"""
        self.send_response(status_code)
        self.send_header('Content-Type', 'application/json')
        self.end_headers()
        self.wfile.write(json.dumps(data).encode('utf-8'))

    def do_POST(self):
        try:
            content_length = int(self.headers.get('Content-Length', 0))
            post_data = self.rfile.read(content_length)
            data = json.loads(post_data.decode('utf-8'))
            
            output_path = data.get('outputPath', 'C:\\n8n-scripts\\output.pptx')
            doc_type = data.get('type', 'presentation')
            
            print(f"[SERVER] Запрошен тип: {doc_type}")
            
            # ==========================================
            # ВЕТКА 1: Генерация WORD-документа
            # ==========================================
            if doc_type == 'word_document':
                word_data_raw = data.get('wordData')
                if isinstance(word_data_raw, str):
                    word_data = json.loads(word_data_raw)
                else:
                    word_data = word_data_raw
                
                script_path = 'C:\\n8n-scripts\\generate_word.py'
                print(f"[SERVER] Запускаем генерацию Word...")
                
                # Запускаем скрипт с таймаутом и перехватом ошибок
                result = subprocess.run(
                    ['python', script_path, json.dumps(word_data), output_path],
                    capture_output=True,
                    text=True,
                    timeout=60 # 60 секунд максимум
                )
                
                if result.returncode != 0:
                    raise Exception(f"Ошибка скрипта: {result.stderr}")
                
                print(f"[SERVER] ✅ Word-документ создан: {output_path}")
                self.send_json_response(200, {"success": True, "filePath": output_path})

            # ==========================================
            # ВЕТКА 2: Генерация ПРЕЗЕНТАЦИИ
            # ==========================================
            else:
                from pptx import Presentation
                from pptx.util import Inches, Pt
                from pptx.dml.color import RGBColor

                pres_data_raw = data.get('presentationData')
                if isinstance(pres_data_raw, str):
                    presentation_data = json.loads(pres_data_raw)
                else:
                    presentation_data = pres_data_raw
                
                prs = Presentation()
                prs.slide_width = Inches(13.333)
                prs.slide_height = Inches(7.5)
                
                slide = prs.slides.add_slide(prs.slide_layouts[0])
                slide.shapes.title.text = presentation_data.get("title", "Презентация")
                slide.shapes.title.text_frame.paragraphs[0].font.size = Pt(44)
                slide.shapes.title.text_frame.paragraphs[0].font.bold = True
                slide.shapes.title.text_frame.paragraphs[0].font.color.rgb = RGBColor(0, 51, 102)
                
                subtitle = slide.placeholders[1]
                subtitle.text = presentation_data.get("subtitle", "")
                subtitle.text_frame.paragraphs[0].font.size = Pt(24)
                
                for slide_data in presentation_data.get("slides", []):
                    slide = prs.slides.add_slide(prs.slide_layouts[1])
                    slide.shapes.title.text = slide_data.get("title", "")
                    slide.shapes.title.text_frame.paragraphs[0].font.size = Pt(36)
                    slide.shapes.title.text_frame.paragraphs[0].font.bold = True
                    slide.shapes.title.text_frame.paragraphs[0].font.color.rgb = RGBColor(0, 51, 102)
                    
                    body_shape = slide.shapes.placeholders[1]
                    tf = body_shape.text_frame
                    tf.clear()
                    
                    bullets = slide_data.get("bullets", [])
                    if bullets:
                        tf.text = bullets[0]
                        tf.paragraphs[0].font.size = Pt(20)
                        for bullet in bullets[1:]:
                            p = tf.add_paragraph()
                            p.text = bullet
                            p.font.size = Pt(20)
                
                prs.save(output_path)
                print(f"[SERVER] ✅ Презентация создана: {output_path}")
                self.send_json_response(200, {"success": True, "filePath": output_path})
            
        except subprocess.TimeoutExpired:
            print("[SERVER ERROR] ❌ Превышено время ожидания (60 сек)")
            self.send_json_response(500, {"error": "Timeout"})
        except Exception as e:
            error_msg = str(e)
            print(f"[SERVER ERROR] ❌ Критическая ошибка: {error_msg}")
            # Записываем полный стек ошибки в файл для отладки
            with open("C:\\n8n-scripts\\server_error.log", "a", encoding="utf-8") as f:
                f.write(traceback.format_exc() + "\n---\n")
            
            # ГАРАНТИРОВАННО отправляем ответ, чтобы n8n не получил HPE_CR_EXPECTED
            self.send_json_response(500, {"error": error_msg})
    
    def do_GET(self):
        try:
            file_path = unquote(self.path[1:])
            if os.path.exists(file_path):
                self.send_response(200)
                if file_path.endswith('.docx'):
                    content_type = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
                else:
                    content_type = 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
                
                self.send_header('Content-Type', content_type)
                self.send_header('Content-Disposition', f'attachment; filename="{os.path.basename(file_path)}"')
                self.end_headers()
                
                with open(file_path, 'rb') as f:
                    self.wfile.write(f.read())
                print(f"[SERVER] ✅ Файл отправлен: {file_path}")
            else:
                self.send_json_response(404, {"error": "File not found"})
        except Exception as e:
            print(f"[SERVER ERROR] ❌ Ошибка отправки файла: {e}")
            self.send_json_response(500, {"error": str(e)})

if __name__ == '__main__':
    server = HTTPServer(('127.0.0.1', 8888), PresentationHandler)
    print("✅ Сервер запущен на http://127.0.0.1:8888")
    print("🛑 Для остановки нажмите Ctrl+C")
    server.serve_forever()