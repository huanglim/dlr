import logging, sys, os
import paramiko

sys.path.append(os.path.dirname(os.getcwd()))
from config import HOSTNAME

logging.basicConfig(level=logging.DEBUG)

def mk_dir(username, hostname):

    # setup a SSHClient object
    ssh = paramiko.SSHClient()

    # 允许将信任的主机自动加入到host_allow 列表，此方法必须放在connect方法的前面
    ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())

    # connect to host
    try:
        ssh.connect(hostname=hostname, port=22, username='seluser', password='seluser')
    except Exception as e:
        raise
    
    # execute the command
    dir_name = '/home/seluser/Downloads/'+ username

    try:
        stdin, stdout, stderr = ssh.exec_command('mkdir '+ dir_name)
    except Exception as e:
        logging.error(e)
    
    # logging the output
    logging.debug('stdout: %s, \nstderr: %s'%(stdout, stderr))
    # close the connection
    ssh.close()

    return dir_name

def trigger_send_to_ftpserver(hostname):
        # setup a SSHClient object
    ssh = paramiko.SSHClient()

    # 允许将信任的主机自动加入到host_allow 列表，此方法必须放在connect方法的前面
    ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())

    # connect to host
    try:
        ssh.connect(hostname=hostname, port=22, username='seluser', password='seluser')
    except Exception as e:
        raise

        # execute the command
    dir_name = '/home/seluser/Downloads/'

    try:
        logging.debug('change dir')
        stdin, stdout, stderr = ssh.exec_command('cd '+ dir_name+'&& pwd')
        for line in stdout:
            print(line)

    except Exception as e:
        logging.error(e)
    else:
        logging.debug('process send to ftp')
        stdin, stdout, stderr = ssh.exec_command('python3 /opt/bin/send_to_ftp.py')
        for line in stdout:
            print(line)

        for line in stderr:
            print(line)
    
    # close the connection
    ssh.close()

if __name__ == '__main__':
    trigger_send_to_ftpserver(HOSTNAME)