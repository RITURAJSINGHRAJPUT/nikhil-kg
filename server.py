import http.server
import json
import os
import base64
import subprocess

class MyHandler(http.server.SimpleHTTPRequestHandler):
    def do_POST(self):
        if self.path == '/api/save':
            content_length = int(self.headers['Content-Length'])
            post_data = self.rfile.read(content_length)
            try:
                payload = json.loads(post_data.decode('utf-8'))
                
                # Check for uploaded image
                img_data = payload.get('image')
                if img_data:
                    filename = img_data.get('filename')
                    b64_content = img_data.get('base64')
                    if filename and b64_content:
                        # Ensure images directory exists
                        os.makedirs('images', exist_ok=True)
                        file_path = os.path.join('images', filename)
                        with open(file_path, 'wb') as img_file:
                            img_file.write(base64.b64decode(b64_content))
                        print(f"Saved uploaded image to: {file_path}")
                
                # Save machines to machines.json
                machines = payload.get('machines')
                if machines is not None:
                    with open('machines.json', 'w', encoding='utf-8') as f:
                        json.dump(machines, f, indent=2, ensure_ascii=False)
                    print("Updated machines.json successfully!")
                
                # Rebuild site
                print("Rebuilding index.html...")
                subprocess.run(['python3', 'build_site.py'], check=True)
                
                self.send_response(200)
                self.send_header('Content-Type', 'application/json')
                self.end_headers()
                self.wfile.write(b'{"status":"success"}')
            except Exception as e:
                print(f"Error handling POST /api/save: {e}")
                self.send_response(500)
                self.send_header('Content-Type', 'application/json')
                self.end_headers()
                self.wfile.write(json.dumps({"error": str(e)}).encode('utf-8'))
        else:
            self.send_response(404)
            self.end_headers()

if __name__ == '__main__':
    PORT = 8000
    server = http.server.HTTPServer(('127.0.0.1', PORT), MyHandler)
    print(f"==================================================")
    print(f" Bookends Hospitality Local Dev Server")
    print(f" Running at: http://127.0.0.1:{PORT}/")
    print(f" Press Ctrl+C to stop the server")
    print(f"==================================================")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nStopping server...")
        server.server_close()
