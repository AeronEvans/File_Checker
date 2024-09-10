import os, paramiko, time



DIR_PATH = os.path.dirname(os.path.abspath(__file__))
INPUT_FOLDER = "Month End Files"
input_files = os.listdir(os.path.join(DIR_PATH,INPUT_FOLDER))


def main():

    host = "sftp.bloomberg.com"
    username = "u108526442"
    password="3bz2YCk0I[qFHAO3"
    port = 22

    with paramiko.SSHClient() as client:
        client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        client.connect(host, port, username, password, disabled_algorithms={'pubkeys': ['rsa-sha2-256', 'rsa-sha2-512']})
        print("Connected")
        with client.open_sftp() as sftp:
            files = sftp.listdir("/report")
            for file in files:
                if "Month_End_Report" in file and "20240903" in file and "xlsx" in file and file not in input_files:
                    print(file)
                    try:
                        sftp.get("/report/"+file, os.path.join(INPUT_FOLDER,file))
                    except:
                        print("error moving file")

    
    




if __name__ == '__main__':
    main()


