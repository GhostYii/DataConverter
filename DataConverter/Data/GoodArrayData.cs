using System;
using System.Collections.Generic;
using GoodArrayData = System.Collections.Generic.List<GoodArray>;

public struct GoodArray
{
    public int id;
    public string name;
    public float ff;
    public TestObject obj_data;
    public List<int> arr;
    public List<List<float>> duoarr;
    public Dictionary<string, int> dict_test;
    public Dictionary<string, Dictionary<string, int>> dict_dict_test;
}
