# 登录与权限

## 概要设计

用于打印按需求格式生成的模板

## 数据库设计

无数据库设计

## 交互接口

注意：对于注册、登录、修改密码这些操作，客户端建议使用HTTPS协议与服务端通信以确保安全。

## 加载插件
	plugin/index.php中加入该插件

```
Plugins::add("print");
```

### 打印

1. 安装第三方打印应用
   \\server-pc\share\software\光速云插件5.0.5.zip

2. 准备打印模板文件, 放置到代码server/template目录下

3. 后端需要支持打印的AC类中, 重载accessControl类里的onHandleExportFormat, 在该函数调用插件
   return printSvc($fmt, $ret, $fname, $tpl);

- fmt: 目前只支持excel, 所以fmt就是"excel"
- ret: 待打印的数据
- fname: 输出文件名
- tpl: 模板文件名称, 默认会在template目录下查找

注意: 要加上这个return, 或者onHandleExportFormat需要return true;

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
1. 待替换的部分原则上用{}包含
2. 字段名对应, 比如查询的字段名是 cmt, 那么模板里定义为{cmt}
3. 支持`.`符号拼接变量, 比如需显示 数量(qty) + 单位(unit), 可使用 {qty.unit}
4. 支持php函数调用, 比如可定义`{date(FMT_D,strtotime(productTm))}`, 会生成生产日期信息
5. 支持子表, 使用`START_ROW::START_SEQ`标识子表起点位置. {子表名}来识别查询结果中的子表, 如START_ROW::START_SEQ{inv1}. 子表只需要定义第一行, 后面会自动循环和填入
6. 支持二维码生成, 需标识IMG和二维码对应的字段, 比如{IMG::qrCodeFlag,itemCode,orderCode,batchNo,cartonNo,qty,bpCode,unit}, 会将这些字段拼接为字符串并生成二维码