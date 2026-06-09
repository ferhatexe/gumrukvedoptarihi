import socket
import threading

def handle_client(client_socket):
    try:
        # Read the initial request headers
        request = client_socket.recv(4096)
        if not request:
            client_socket.close()
            return
        
        # Parse the connection request
        first_line = request.decode('latin1').split('\n')[0]
        words = first_line.split()
        if len(words) < 2:
            client_socket.close()
            return
            
        method, url = words[0], words[1]
        
        if method == 'CONNECT':
            # HTTPS tunnel (used by secure customs connection)
            try:
                if ':' in url:
                    host, port = url.split(':')
                    port = int(port)
                else:
                    host = url
                    port = 443
            except Exception:
                client_socket.close()
                return
            
            # Connect to the target customs website
            remote_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            remote_socket.connect((host, port))
            
            # Send connection established status response
            client_socket.sendall(b"HTTP/1.1 200 Connection Established\r\n\r\n")
            
            # Pipe data bidirectionally between Render and the customs site
            def pipe(src, dst):
                try:
                    while True:
                        data = src.recv(4096)
                        if not data:
                            break
                        dst.sendall(data)
                except Exception:
                    pass
                finally:
                    src.close()
                    dst.close()
            
            t1 = threading.Thread(target=pipe, args=(client_socket, remote_socket))
            t2 = threading.Thread(target=pipe, args=(remote_socket, client_socket))
            t1.start()
            t2.start()
            
        else:
            # Standard HTTP request forward
            host_header = ""
            for line in request.decode('latin1').split('\n'):
                if line.lower().startswith('host:'):
                    host_header = line.split(':')[1].strip()
                    break
            if not host_header:
                client_socket.close()
                return
            
            if ':' in host_header:
                host, port = host_header.split(':')
                port = int(port)
            else:
                host = host_header
                port = 80
                
            remote_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            remote_socket.connect((host, port))
            remote_socket.sendall(request)
            
            # Read response and send back to client
            while True:
                data = remote_socket.recv(4096)
                if not data:
                    break
                client_socket.sendall(data)
            remote_socket.close()
            client_socket.close()
            
    except Exception:
        try:
            client_socket.close()
        except Exception:
            pass

def main():
    server = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    server.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
    # Bind to local port 8888
    server.bind(('0.0.0.0', 8888))
    server.listen(100)
    print("--------------------------------------------------")
    print(" İş Yeri Bağlantı Tüneli (Proxy) Başlatıldı")
    print(" Port: 8888")
    print("--------------------------------------------------")
    while True:
        client_sock, addr = server.accept()
        t = threading.Thread(target=handle_client, args=(client_sock,))
        t.daemon = True
        t.start()

if __name__ == '__main__':
    main()
