# fillgooglesheets
把自动填google表格封装成接口
通过curl发送post请求调用接口
自动插入传参到对应的表格
insert.py里面为填表逻辑
taw.json里面写入了对应的参数和表格的关系
app.py为封装成flask的代码
gstart.sh为添加wsgi启动多线程保证高可用
