Sub 宏1()
Attribute 宏1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 宏1 宏
'

'将源列K列每行数据复制成4份顺序放入目标列P列
    Dim i
    Dim k
    Dim p
    Dim k_end
'k复原源列开始行
    k = 2
'k_end复制源列终行
    k_end = 10
'i索引
    i = 0
'p复制目标列开始行
    p = 2
    Do While (k < k_end)
        i = 0
        Do While (i < 4)
'"p"P列（目标列），"k"K列（源列）
            Range("p" & p) = Range("k" & k)
            p = p + 1
            i = i + 1
        Loop
        k = k + 1
    Loop
    
End Sub
