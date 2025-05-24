using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace VisioReader
{
    /// <summary>
    /// Custom JSON converter that skips serializing empty collections.
    /// When a collection is empty, this converter will not include it in the JSON output.
    /// </summary>
    public class EmptyCollectionConverter : JsonConverterFactory
    {
        public override bool CanConvert(Type typeToConvert)
        {
            // Check if the type is an enumerable but not a string
            return typeof(IEnumerable).IsAssignableFrom(typeToConvert) && 
                   typeToConvert != typeof(string);
        }

        public override JsonConverter CreateConverter(Type typeToConvert, JsonSerializerOptions options)
        {
            // Use reflection to create the appropriate converter type for the collection
            var converterType = typeof(EmptyCollectionConverterInner<>).MakeGenericType(typeToConvert);
            return (JsonConverter)Activator.CreateInstance(converterType)!;
        }
          private class EmptyCollectionConverterInner<T> : JsonConverter<T> where T : IEnumerable
        {
            public override T? Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
            {
                // Create a clone of options without this converter to avoid infinite recursion
                var clonedOptions = new JsonSerializerOptions(options);
                var convertersToKeep = new List<JsonConverter>();
                foreach (var converter in clonedOptions.Converters)
                {
                    if (!(converter is EmptyCollectionConverter))
                    {
                        convertersToKeep.Add(converter);
                    }
                }
                
                clonedOptions.Converters.Clear();
                foreach (var converter in convertersToKeep)
                {
                    clonedOptions.Converters.Add(converter);
                }
                
                // For reading, use the default JSON deserializer
                return JsonSerializer.Deserialize<T>(ref reader, clonedOptions);
            }

            public override void Write(Utf8JsonWriter writer, T value, JsonSerializerOptions options)
            {
                // Skip writing if the collection is empty
                if (value == null || !value.GetEnumerator().MoveNext())
                {
                    // Don't write anything for empty collections
                    return;
                }

                // Create a clone of options without this converter to avoid infinite recursion
                var clonedOptions = new JsonSerializerOptions(options);
                var convertersToKeep = new List<JsonConverter>();
                foreach (var converter in clonedOptions.Converters)
                {
                    if (!(converter is EmptyCollectionConverter))
                    {
                        convertersToKeep.Add(converter);
                    }
                }
                
                clonedOptions.Converters.Clear();
                foreach (var converter in convertersToKeep)
                {
                    clonedOptions.Converters.Add(converter);
                }
                
                // For non-empty collections, use the default serializer
                JsonSerializer.Serialize(writer, value, clonedOptions);
            }
        }
    }
}
