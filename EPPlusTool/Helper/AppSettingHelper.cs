using System;
using System.Collections.Generic;
using System.Linq;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace EPPlusTool.Helper
{
    public class AppSettingHelper<T> where T : class, new()
    {
        /// <summary>
        /// 程序没法确保section的json  就是对应的模型
        /// </summary>
        /// <param name="physicalPath"></param>
        /// <param name="section">不要写错, 写错会造成整个文件的数据不对,而不是section的这一块数据不错,如果是根目录, 可以传入"" 或null </param>
        /// <param name="applyChanges"></param>
        /// <returns></returns>
        public static bool SetAppSettingValue(string physicalPath, string section, Action<T> applyChanges)
        {
            JObject jObject = AppSettingHelperCommon.GetJObject(physicalPath);

            JToken jToken;

            if (string.IsNullOrEmpty(section))
            {
                if (jObject != null && jObject is JToken)
                {
                    jToken = jObject as JToken;
                }
                else
                {
                    throw new ArgumentException(nameof(section));
                }
            }
            else
            {
                var (haveJToken, jObjectKey2, jToken2) = AppSettingHelperCommon.GetJTokenFlatSection(jObject, section);
                if (!haveJToken)
                {
                    return false; //没有对应的section
                }
                else
                {
                    section = jObjectKey2;

                    jToken = jToken2;
                }
            }

            if (jToken is null)
            {
                return false; //找不到对应的值
            }

            var sectionObject = JsonConvert.DeserializeObject<T>(jToken.ToString());
            applyChanges(sectionObject);

            var value = JObject.Parse(JsonConvert.SerializeObject(sectionObject));

            //jToken = value;

            if (string.IsNullOrEmpty(section))
            {
                jObject = value;
            }
            else
            {
                jObject[section] = value;
            }
            AppSettingHelperCommon.Save(physicalPath, jObject);

            return true;
        }

    }

    public class AppSettingHelper
    {

        /// <summary>
        /// 
        /// </summary>
        /// <param name="physicalPath"></param>
        /// <param name="section">不区分大小写的(注:Newtonsoft.json 是区分的)</param>
        /// <param name="value"></param>
        public static bool SetAppSettingValue(string physicalPath, string section, object value)
        {

            JObject jObject = AppSettingHelperCommon.GetJObject(physicalPath);

            var (haveJToken, jObjectKey, jToken) = AppSettingHelperCommon.GetJTokenFlatSection(jObject, section);

            if (!haveJToken)
            {
                return false;
            }

            if (jToken is JValue jTokenValueIsJValue)
            {
                //if (jTokenValueIsJValue.Value != value)//前者类型是object(xx) 后者是xx ,比较永远false, 改用object.Equals
                if (value is int)
                {
                    if (object.Equals(value, jTokenValueIsJValue.Value))
                    {
                        return true;
                    }

                    if (object.Equals(Convert.ToInt64(value), jTokenValueIsJValue.Value))
                    {
                        return true;
                    }
                }

                if (object.Equals(value, jTokenValueIsJValue.Value))
                {
                    return true;
                }

                jTokenValueIsJValue.Value = value;
                AppSettingHelperCommon.Save(physicalPath, jObject);
                return true;
            }

            return false;
        }

    }

    public class AppSettingHelperCommon
    {
        /// <summary>
        /// 扁平化的section ,seciton 不区分大小写
        /// </summary>
        /// <param name="jObject"></param>
        /// <param name="section">扁平化的section , 如a:b:c</param>
        /// <returns></returns>
        public static (bool haveJToken, string jObjectKey, JToken jToken) GetJTokenFlatSection(JObject jObject, string section)
        {

            var sectionArray = section.Split(new[] { ':' }, StringSplitOptions.RemoveEmptyEntries);

            {
                if (sectionArray.Length <= 0)
                {
                    return (false, null, null);
                }
            }

            {
                if (sectionArray.Length == 1)
                {
                    var (hasJObjectKey, jObjectKey, jTokenValue) = AppSettingHelperCommon.GetJTokenSingSection(jObject, sectionArray[0]);

                    if (hasJObjectKey)
                    {
                        return (true, jObjectKey, jTokenValue);
                    }
                    return (false, null, null);
                }
            }

            {
                //else sectionArray.Length>1

                var a = AppSettingHelperCommon.GetJTokenSingSection(jObject, sectionArray[0]);
                if (!a.hasJObjectKey)
                {
                    return (false, null, null);
                }

                for (var index = 1; index < sectionArray.Length; index++)
                {
                    if (a.JTokenValue is JObject jTokenValueAsJObject)
                    {
                        a = AppSettingHelperCommon.GetJTokenSingSection(jTokenValueAsJObject, sectionArray[index]);
                        if (!a.hasJObjectKey)
                        {
                            return (false, null, null);
                        }
                    }
                    else
                    {
                        return (false, null, null);
                    }
                }
                return (a.hasJObjectKey, a.JObjectKey, a.JTokenValue);
            }
        }


        /// <summary>
        /// 获得不区分大小写的key 所对应的key.  JObject的key 是区分大小写的, 
        /// </summary>
        /// <param name="jObject"></param>
        /// <param name="section">单个section ,不能有嵌套(a:b这种的),有嵌套就找不到</param>
        /// <returns></returns>
        public static (bool hasJObjectKey, string JObjectKey, JToken? JTokenValue) GetJTokenSingSection(JObject jObject, string section)
        {
            //先按区分大小写找, 找不到按不区分大小写找
            {
                var exists = jObject.TryGetValue(section, out var jTokenValue);
                if (exists)
                {
                    return (true, section, jTokenValue);
                }
            }


            var dict = new Dictionary<string, string>();// key: JObject的key.Tolower()  value  JObject的key

            var matchedPattern = 1;// 不区分大小写

            foreach (var currentObj in jObject)
            {
                var dictKey = currentObj.Key.ToLower();
                if (dict.ContainsKey(dictKey))
                {
                    matchedPattern = 2;//区分大小写;
                }
                else
                {
                    var dictValue = currentObj.Key;
                    dict.Add(dictKey, dictValue);
                }
            }

            if (matchedPattern == 1)
            {
                var lowerSection = section.ToLower();
                if (dict.ContainsKey(lowerSection))
                {
                    return (true, dict[lowerSection], jObject.GetValue(dict[lowerSection]));
                }
                else
                {
                    return (false, section, null);
                }
            }
            //这段逻辑放在最开头了
            //if (matchedPattern == 2)
            //{
            //var exists = jObject.TryGetValue(section, out JToken? JTokenValue);
            //if (exists)
            //{
            //    return (true, section, JTokenValue);
            //}
            //else
            //{
            //    return (false, section, null);
            //}
            //}
            return (false, section, null);

        }

        public static JObject GetJObject(string physicalPath)
        {
            if (physicalPath is null)
            {
                physicalPath = System.IO.Path.Combine(System.AppContext.BaseDirectory, "appsettings.json");
            }

            var jsonPara = System.IO.File.ReadAllText(physicalPath);
            JObject jsonObj = JsonConvert.DeserializeObject<JObject>(jsonPara);
            return jsonObj;
        }

        public static void Save(string physicalPath, JObject jsonObj)
        {
            string output = JsonConvert.SerializeObject(jsonObj, Formatting.Indented);
            System.IO.File.WriteAllText(physicalPath, output);
        }

    }
}
