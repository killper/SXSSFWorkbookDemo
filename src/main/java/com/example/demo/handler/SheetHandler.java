package com.example.demo.handler;

import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

/**
 * @author lucasliang
 * @date 09/01/2019 11:11 上午
 * @Description
 */
public class SheetHandler extends DefaultHandler {

  /**
   * 解析到XML的开始标签触发此方法
   *
   * @param uri 如"http://schemas.openxmlformats.org/spreadsheetml/2006/main"
   * @param localName 当前开始标签的名字
   * @param name 标签名
   * @param attributes 当前标签的属性对象
   */
  @Override
  public void startElement(String uri, String localName, String name, Attributes attributes)
      throws SAXException {
    if ("row".equals(name)) {// <row>:开始处理某一行
      //
    } else if ("c".equals(name)) {// <c>:一个单元格
      //
    } /*else if (isTextTag(name)) {// <v>:单元格值
      //
    }*/ else if ("f".equals(name)) {// <f>:公式表达式标签
      //
    } else if ("is".equals(name)) {// 内联字符串外部标签
      //
    } else if ("col".equals(name)) {// 处理隐藏列
      //
    }
  }

  /**
   * 返回单元格的值
   */
  @Override
  public void characters(char[] ch, int start, int length) throws SAXException {
    StringBuffer sb = new StringBuffer();
    sb.append(ch, start, length);
  }

  /**
   * 解析到XML的结束标签触发此方法 如：</row>
   */
  @Override
  public void endElement(String uri, String localName, String name) throws SAXException {
    if ("c".equals(name)) {// <c>标签结束
      //
    }/* else if (isTextTag(name)) {// 文本单元格结束标签
      //
    }*/ else if ("row".equals(name)) {// 行结束标签
      //
    } else if ("f".equals(name)) {// </f>标签
      //
    } else if ("is".equals(name)) {// </is>标签
      //
    } else if ("worksheet".equals(name)) {// Sheet读取完成
      //
    }
  }
}
