# -*- coding: UTF-8 -*-

'''
Created on 2013-4-8

@author: simon_000
'''
import os;
import sys;
import boto.sqs;
import boto.sns;

import json;

import win32com.client;
import win32com.gen_py.MSO as MSO;
import win32com.gen_py.PO as PO;
from boto.s3.connection import S3Connection
from boto.s3.key import Key
from boto.sqs.message import Message

#有关AWS服务的配置
TOKYO_REGION = 'ap-northeast-1';
LBW_AWS_ACCESS_KEY = 'AKIAIEROLA5Y34CYM6TA';
LBW_AWS_SECRET_KEY= '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';

QUEUE_NAME = "LivePPT-pptId-Bus";
BUCKET_NAME = "pptstore";
TOPIC_ARN = "arn:aws:sns:ap-northeast-1:206956461838:liveppt-sns";

MAX_QUEUE_WAIT_TIME = 20;

UTF8_ENCODING = "UTF-8";

g = globals()
for c in dir(MSO.constants): g[c] = getattr(MSO.constants, c) # globally define these
for c in dir(PO.constants): g[c] = getattr(PO.constants, c)

#准备S3服务
S3_Conn = S3Connection(LBW_AWS_ACCESS_KEY, LBW_AWS_SECRET_KEY);
pptstore_bucket = S3_Conn.get_bucket(BUCKET_NAME);
k = Key(pptstore_bucket);




#准备SQS服务
SQS_Conn = boto.sqs.connect_to_region(TOKYO_REGION,\
                aws_access_key_id = LBW_AWS_ACCESS_KEY,\
                aws_secret_access_key = LBW_AWS_SECRET_KEY);
                
q = SQS_Conn.create_queue(QUEUE_NAME);

#准备SNS服务
SNS_Conn = boto.sns.connect_to_region(TOKYO_REGION,\
                aws_access_key_id = LBW_AWS_ACCESS_KEY,\
                aws_secret_access_key = LBW_AWS_SECRET_KEY);
                

# 准备PPT应用
Application = win32com.client.Dispatch("PowerPoint.Application")
Application.Visible = True

print sys.getdefaultencoding();
print u'LivePPT-PPT-Converter is launched！';

# 测试用途
# ppt_id = 'eb5697be-9ff0-467c-a7df-36def1ac9001';
# mm = Message();
# mm.set_body(ppt_id.encode(UTF8_ENCODING));
# q.write(mm);

ppt_dir_path="I:\\ppt";

while True:
    m = q.read(wait_time_seconds = MAX_QUEUE_WAIT_TIME);
    if m<>None:
        pptId = m.get_body().encode(UTF8_ENCODING);
        q.delete_message(m);
        print pptId;
        print '\n';
        
        #准备路径参数
        ppt_path = ppt_dir_path+"\\"+pptId; #PPT存放位置
        save_dir_path = ppt_dir_path+"\\converted_"+pptId; #保存转换后图片的文件夹路径
        
        #从S3获取文件并存入本地
        k.key = pptId;
        f = file(ppt_path,"wb");
        k.get_file(f);
        f.close();
        
        #使用PowerPoint打开本地PPT，并进行转换
        try:            
            myPresentation = Application.Presentations.Open(ppt_path);
            myPresentation.SaveAs(save_dir_path, ppSaveAsPNG);
        finally:
            myPresentation.Close();
            
        png_file_name_list = os.listdir(save_dir_path);
        ppt_count = len(png_file_name_list);
        for i in range(1, ppt_count):
            png_path = save_dir_path+u"\\幻灯片"+str(i)+u".PNG"; #单个PNG文件路径
            png_key = pptId + "-"+ str(i);
            print i;
            print png_path;
            
            #上传单个文件
            try:
                f = open(png_path,"rb");
                k.key = png_key;
                k.set_contents_from_file(f);
            finally:
                f.close();
        
        #组装准备发到SNS的信息
        sns_message = {};
        sns_message['isSuccess'] = True;
        sns_message['pptId']=pptId;
        sns_message['count'] = "ppt_count";
        
        #发送消息到SNS
        SNS_Conn.publish(TOPIC_ARN, json.dumps(sns_message));
        
#         Application.Quit();
        
         
        
