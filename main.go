package main

import (
    "errors"
    "flag"
    "fmt"

)

var filePath string
func init() {
    flag.StringVar(&filePath, "filePath", "", "请输入文件路径")

    flag.Parse()
}

func check() error {
    if filePath == "" {
        return errors.New("文件路径不能为空")
    }

    fmt.Println("filePath:", filePath)

    return nil
}

func main() {
    // 参数验证
    if err := check(); err != nil {
        fmt.Println("参数校验失败，"+err.Error())
        return
    }

    source := "/Users/fanfei/Downloads/test1.doc"
    destination := "/Users/fanfei/Downloads/test1.xml"

    err := Doc2XML(source, destination)
    if err != nil {
        fmt.Println("文件转换失败，"+err.Error())
        return
    }

    fmt.Println("文件转换成功")
    return
}
