#excel2json


## Overview

通过Python的xlrd库, 导出excel的数据到Json文件.

## Usage

+ 直接通过Python的运行环境执行, 需要安装xlrd  

```
pip install xlrd
python excel2json
```

+ 通过py2exe打包成exe文件.

```
cd dist
./excel2json.exe
```

## Setting

config.json:
```
{
	"work_dir"   : ".",
	"output_dir" : "./output"
}
```
