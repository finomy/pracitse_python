import os
import shutil
import pymysql
import time


def copy_file(source, dst):
    for file_name in os.listdir(source):
        print(file_name)
        if '.txt' in file_name:
            file_full_path = os.path.join(source, file_name)
            dst_full_path = os.path.join(dst, file_name)
            print(file_full_path, dst_full_path)
            shutil.copyfile(file_full_path, dst_full_path)


def clear_file_with_tag(path, tag):
    for file_name in os.listdir(path):
        if tag in file_name:
            file_full_path = os.path.join(path, file_name)
            os.remove(file_full_path)
            if file_name in os.listdir(path):
                print('delete failed', file_name)
            else:
                print('delete successed', file_name)


def count_file_with_tag(path, tag):
    num = 0
    for file_name in os.listdir(path):
        if tag in file_name:
            num += 1
    print(num)
    return num


def clear_db_table(mysql_ip, username, psw, db_name):
    db = pymysql.connect(mysql_ip, username, psw, db_name)
    cursor = db.cursor()
    cursor.execute('DELETE FROM ')
    db.commit()
    cursor.execute('DELETE FROM ')
    db.commit()
    cursor.execute('DELETE FROM ')
    db.commit()
    db.close()


def delete_table(mysql_ip, username, psw, db_name):
    db = pymysql.connect(mysql_ip, username, psw, db_name)
    cursor = db.cursor()
#    cursor.execute("select table_name from information_schema.tables where table_name like '%info' "
#                   "and table_schema = 'qfj_clearing_box'")
    cursor.execute("SELECT table_name FROM information_schema.tables WHERE table_name like '%3' and table_schema = 'testdb1'")
    info_tables = cursor.fetchall()
#    cursor.execute("select table_name from information_schema.tables where table_name like '%image' "
#                   "and table_schema = 'qfj_clearing_box'")
#    image_tables = cursor.fetchall()
    for item in info_tables:
        print(item[0])
        cursor.execute("DROP TABLE "+(item[0]))
        db.commit()
    cursor.execute('DELETE FROM t1')
    db.commit()
#    print(image_tables)
    db.close()


# copy_file('D:\\battest', 'D:\\battest\\t')
# count_file_with_tag('D:\\battest\\t', '.txt')
# clear_file_with+tag('D:\\battest\\t', '.txt')
# count_file_with_tag('D:\\battest\\t', '.txt')

# delete_table('localhost', 'root', '', 'testdb1')

