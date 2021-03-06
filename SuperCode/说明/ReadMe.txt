这个文件简述SuperCode的密码体制。

SuperCode可以对任意文件进行Encript操作，生成加密文件；对加密文件进行Decript操作，生成明文文件。

加解密操作的单位是字节，属于流密码。密钥是一个不断变化的散裂函数，详细介绍见下文。
==============
SuperCode的加密详细步骤：

加密者将遍历明文的每个字节。

首先，加密者拥有一个加密密钥Key。key是一个表，展示单字节明文密文的转换关系。
一个Key的例子：
（在这个例子中，为了方便Byte取8而非256）

明文	密文
0	4
1	2
2	1
3	6
4	3
5	5
6	0
7	7

Key最好是随机生成。
有了Key，加密者用这个Key加密明文的第一个字节，写入密文的第一个字节。
接着，加密者再次随机生成一个Key2，并把Key2以Key作为密钥加密，并写入密文的第2~257个字节。
然后，加密者用Key2作为密钥加密明文第二个字节，写入密文，在生成Key3，用Key2加密，依次列推...

===============
SuperCode的解密方式：

解密者拥有原始密钥Key，成功解密密文的第1~257个字符。这样，他就获得了第一个明文字符以及Key2。依次列推，他迭代获得全部的密钥和明文。

=============
SuperCode的优势：

一个密文之所以能在密钥不泄露的情况下被第三方破解，是因为密文泄露了信息。
就举位移密码这个例子吧。位移密码极易破解，一个原因是英文中e的出现频率最高，而密文中不同字符的出现频率有高有低，从而为破译者提供了信息。这个例子中，“密文各个字符出现频率的不随机性”给破译者提供信息。
在更加高级的密码体系中，虽然“密文各个字符出现的频率”相等了，但是在更隐晦的层面上，密文还是给破译者暴露了信息。
所以，得出结论，只有当一段密文不论在什么层面上都显现出完全的随机性，那么这个密文才是不可在不知道密钥的情况下破译的。

接着，我们必须意识到，好的加密体制本质上是在“掩盖语言本身的缺陷”。何谓缺陷？拿英语来说吧，各个字母出现的频率不一样，就是一个缺陷。设想如果英语本身各个字母出现的频率就一样，用位移加密之后的密文当然不会出现个个字符出现频率不同的问题。
那么，存不存在“完美的语言”呢？它又是什么样子的呢？
进过一番思考，不难得出结论，如果存在一种完美的语言，那么在这种语言中，任何一串字符序列必定都是有意义的，随机序列也不例外。即，baoiw7eh4jalweblja在那种语言中一定是有意义的。为什么这么说呢？设想有一篇用这种完美语言撰写的明文，就算加密者用最简单的位移方法加密，得到的密文也是难以破解的——破译者若用26个可能的密钥去枚举，他将得到26个语句通顺、意思不同的明文版本，而不知道那个才是真的。如果破译者不知道加密者是用位移加密的，那就更复杂了，因为任意 f(密文) 都是语句通顺且有意义的。可以断言，破译者无法在没有密钥的情况下成功破译完美语言下的密文。

我们说，这种完美的语言存在“语言的全对称性”。它不存在一个自相矛盾或是不通顺的语段，让破译者有机会排除掉任何一个密钥的可能。

但是很可惜，这种语言并不存在在。

那么，有没有一种加密体制能够把密文伪装成“完全对称”的呢？有没有方法能让密文看起来和随机序列丝毫没有差别呢？

“流密码”是一个这样的方法。它的密钥是一串与明文等长的随机二进制数列。
加密方法是：
密文=明文xor密钥
其中，xor是异或门。

稍加思考就会发现，只要密钥的随机性够好，生成的密文就应该是看上去完全随机的。破译者如果遍历所有可能的密钥，将会得到宇宙里所有用英文能表达的意思，而且他还不能辨别这些意思中哪些意思是正真明文的概率更大些。

然而，这个流密码的缺陷也显而易见——使用双方在分手之前必须约定一串很长很长的密钥，同时还要保证密钥的随机性良好。因为密钥的瑕疵，最终导致的就是密文的瑕疵。

SuperCode的优势就显现了。SuperCode中，双方需要约定的密钥只有原始密钥Key，长度只有256字节。使用者甚至可以把它抄在一张小纸片上随身带着。

===========
SuperCode的劣势：
尽管Supercode被设计成尽量让密文显得完全随机，但是我目前还不能证明SuperCode生成的密文具有流密码具有的“全对称性”。也许它存在一个漏洞，让它在某个层面上失去完美的随机性，从而变得可能在无密钥情况下被破译。

SuperCode的另一点劣势也显而易见，就是它生成的密文要比原文大256倍。。。最然密文大比密钥大(流密码)让人舒服多了，但是密文这么长又无法被压缩(如果满足完全对称性，被压缩是不可能的)也是一件头疼的事情。

The End
==============
SuperCode.exe使用说明
文本框里输入操作路径(Root)
Encription将会加密Root\text.txt并把密文写进Root\code
Decription将会解密Root\code并把明文写进Root\text.txt
Create将会生成随机密钥Root\key

此程序支持对三个不同Root的操作。
