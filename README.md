# 登录与权限

## 概要设计

用于打印按需求格式生成的模板

## 数据库设计

无数据库设计

## 交互接口

注意：对于注册、登录、修改密码这些操作，客户端建议使用HTTPS协议与服务端通信以确保安全。

## 加载插件
参见插件安装方式(http://oliveche.com/jdcloud-site/后端框架.html)

+ 用git下载插件，注意与项目目录平级
```
git clone server-pc:src/jdcloud-plugin-print
cd myproject
```

+ 安装
```
./tool/jdcloud-plugin add ../jdcloud-plugin-print
```

### 打印

1. 安装第三方打印应用
	\\server-pc\share\software\光速云插件5.1.3-win7.exe

2. 准备打印模板文件, 放置到代码server/template目录下

3. 后端需要支持打印的AC类中, 加入use exportUtil声明

4. 前端调用

调用插件getShareUrl获取打印地址:

getShareUrl(ac, cond, fmt, fname, tpl)

- fmt: 目前只支持excel, 所以fmt就是"excel" 
- ret: 待打印的数据
- fname: 输出文件名
- tpl: 模板文件名称, 默认会在template目录下查找

```
WUI.callSvr("getShareUrl", {ac:"MaterialTag.query",cond:"id=" + row.id, fmt:"excel", fname:"test", tpl:"tag"},function(res){
				app_show('url: ' + res);
			}{printType:1})
		})
```

调用插件GSCloudPlugin.GetPrinters获取打印机
```
GSCloudPlugin.GetPrinters({
		OnSuccess:function(result){
			printExcel(url,result.Data[0])
		},
		OnError:function(result){
			outputResult(result);
		}
	});
```

调用插件GSCloudPlugin.PrintExcel打印
```
	GSCloudPlugin.PrintExcel({
		Title:"Excel0001", //主题；如果不需要，可以不传此参数
		Url: url, // 支持URL地址与Base64
		PrinterName:printerName,
		OnSuccess:function(result){
			app_show('打印成功')
		},
		OnError:function(result){
			app_show('打印失败')
		}
	});
}
```

## 语法支持
+ 待处理的关键字或变量用`{}`包含

+ 模板内字段名与查询返回的字段名对应, 比如查询的字段名是`cmt`, 那么模板里定义为`{cmt}`

+ 支持变量和常量混用, 比如同一单元格内填入`{qty}件`, 会被替换为具体的数量, 比如 `10件`.

+ 同一单元格支持多变量, 比如`{qty}{unit}`, 会被替换为 `10件`.

+ 多个变量拼接. 可以使用`.`符号拼接变量, 比如需显示`qty` + `unit`, 可使用 `{qty.unit}` 

+ 支持php函数调用, 比如可定义`{date(FMT_D,strtotime(productTm))}`, 会生成生产日期信息

+ 支持for循环输出主表或者子表多行记录, 循环的起点使用`{j_for}`标识, 如果是子表, 那么使用`{j_for::子表名}`.多行输出只需要定义第一行, 插件会根据查询返回的行数自动填充行数
比如:	
	```
	{j_for::inv1}{id}
	```
	标识该行是inv1子表的起点, 本列填入inv1子表的id字段

	如果是多行主表,不写子表名即可, 比如
	```
	{j_for}{orderNo}
	```
	+ 对多行的情况支持自增加序号函数`autoNo()`, 比如首列需要填入从1开始的自增长行号:

		```
		{j_for}{autoNo()}
		```

+ 图片插入. 使用`{j_img}`标识. 格式为`{j_img}{imgPath;width?;height?}`, 比如插入一张地址为`http://oliveche.com/print/1.jpg`的图片
	```
	{j_img}{'http://oliveche.com/print/1.jpg';120;120}
	```
	类似于`<img src=xxxx.jpg>`的用法.

	+ 网络地址需要用`http/https`开头
	+ 如果是无需处理的常量字符串, 使用单引号 `'` 引起来
	+ 宽和高直接在模板里指定. 用分号分隔, 比如示例里的`120`. 如果不指定, 默认图片会是`75*75`的尺寸. 建议自行制定宽高

+ 二维码生成
使用`qrcode(content)`函数生成二维码, 比如

	```
	{j_img}{qrcode(oliveche.com/jd_cloud/api/order.get?id=1)}
	```
	可以生成对应url的一张二维码

+ 二维码字段拼接: 
使用`concatField(...args)`拼接字段
二维码生成一种常见的情况是,使用查询的字段名和值拼接成一段字段串, 生成二维码, 比如想把itemCode,batchNo,cartonNo这三个字段及对应的值拼接成字符串,那么按如下实现:
	```
	{j_img}{qrcode(concatField(itemCode,batchNo,cartonNo))}
	```

+ 图片内容从一个变量获取:
使用`getField()`函数. 比如查询返回`url`字段. 图片地址就是这个`url`字段的值,那么模板里可以如下定义:
```
{j_img}{getField(url)}
```
