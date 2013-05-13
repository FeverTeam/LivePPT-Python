# -*- coding: UTF-8 -*-

'''
Created on 2013-4-8

@author: simon_000
'''
from __future__ import with_statement

#导入
import config as conf
import prepare
import path_utils as p

import os
import sys
import logging
import json

import win32com.client
import win32com.gen_py.MSO as MSO
import win32com.gen_py.PO as PO

import boto.sqs
import boto.sns
from boto.s3.connection import S3Connection
from boto.s3.key import Key



def main():
    g = globals()
    for c in dir(MSO.constants): g[c] = getattr(MSO.constants, c) # globally define these
    for c in dir(PO.constants): g[c] = getattr(PO.constants, c)

    #准备S3服务
    S3_Conn, pptstore_bucket, k = prepare.gen_s3();

    #准备SQS服务
    SQS_Conn, q = prepare.gen_sqs()

    #准备SNS服务
    SNS_Conn = prepare.gen_sns()

    # 准备PPT应用
    Application = win32com.client.Dispatch(conf.POWERPOINT_APPLICATION_NAME)
    Application.Visible = True


    reload(sys)
    sys.setdefaultencoding(conf.UTF8_ENCODING)
    print sys.getdefaultencoding()
    print 'LivePPT-PPT-Converter is launched.'

    # 测试用途
    # ppt_id = 'eb5697be-9ff0-467c-a7df-36def1ac9001'
    # mm = Message()
    # mm.set_body(ppt_id.encode(UTF8_ENCODING))
    # q.write(mm)

    while True:
        m = q.read(wait_time_seconds = conf.MAX_QUEUE_WAIT_TIME)
        #若消息不为空
        if m<>None:
            ppt_id = m.get_body_encoded()
            q.delete_message(m)      
            print ppt_id
            
            #准备路径参数
            ppt_path = p.gen_ppt_path(ppt_id) #PPT存放位置
            save_dir_path = p.gen_save_dir_path(ppt_id) #保存转换后图片的文件夹路径
            
            #从S3获取文件并存入本地
            k.key = ppt_id
            with open(ppt_path, "wb") as f:
                k.get_file(f)
                
            #使用PowerPoint打开本地PPT，并进行转换
            try:            
                myPresentation = Application.Presentations.Open(ppt_path)
                myPresentation.SaveAs(save_dir_path, ppSaveAsJPG)
            finally:
                myPresentation.Close()
                
            png_file_name_list = os.listdir(save_dir_path)
            ppt_count = len(png_file_name_list)
            print 'ppt_page_count'+str(ppt_count)
            for index in range(1, ppt_count+1):
                png_path = p.gen_single_png_path(ppt_id, index) #单个PNG文件路径
                png_key = ppt_id + "p"+ str(index)
                print "uploading " + str(index)
                # print png_path
                
                #上传单个文件
                with open(png_path, "rb") as f:
                    k.key = png_key
                    k.set_contents_from_file(f)
            
            #组装准备发到SNS的信息
            sns_message = {}
            sns_message['isSuccess'] = True
            sns_message['storeKey'] = ppt_id
            sns_message['pageCount'] = ppt_count
            
            #发送消息到SNS
            SNS_Conn.publish(conf.TOPIC_ARN, json.dumps(sns_message))
            
    #         Application.Quit()
    return

if __name__ == '__main__':
    main()