internal static class Program
{
    private static async Task Main()
    {
        var testArray = new [] { 1, 3, 5, 7  };
        Console.WriteLine(Getindex(testArray, 4));
    }

    private static int Getindex(int[] inputArray, int number)
    {private static int Getindex(int[] inputArray, int number)
    {
        //// the inputArray is sort by asc
        //// return the index that the number should be in the order of array
        
        var arrayLength = inputArray.Length;
        
        // 如果数字小于第一个元素，应该插入到索引0
        if (number <= inputArray[0])
            return 0;
        
        // 如果数字大于最后一个元素，应该插入到数组末尾
        if (number > inputArray[arrayLength - 1])
            return arrayLength;
        
        for (int i = 0; i < arrayLength - 1; i++)
        {
            var nextIndex = i + 1;
            
            if (inputArray[i] < number && inputArray[nextIndex] >= number)
            {
                return nextIndex;
            }
        }
        
        return arrayLength;
    }
        //// the inputArray is sort by ace
        //// return the index that the number should be in the order of arrary
        
        
        var arrayLength = inputArray.Length;
        for (int i = 0; i < arrayLength; i++)
        {
            var nextIndex = i++;
            //// next number is the last 
            
            // var testArray = new [] { 1, 3, 5, 7  }; 4
            if (inputArray[i] < number && inputArray[nextIndex] >= number  )
            {
                return i;
            }
        }
        
        return 0;
    }  
    
}