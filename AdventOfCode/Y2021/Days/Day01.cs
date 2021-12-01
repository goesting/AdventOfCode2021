using AdventOfCode.Helpers;

namespace AdventOfCode.Y2021.Days
{
    public static class Day01
    {
        static int day = 1;
        static List<int>? inputs;

        public static string? Answer1 { get; set; }
        public static string? Answer2 { get; set; }

        public static void Run(int part, bool test)
        {
            inputs = InputManager.GetInputAsInts(day, test);

            var start = DateTime.Now;

            string part1 = "";
            string part2 = "";

            switch (part)
            {
                case 1:
                    part1 = Part1();
                    break;
                case 2:
                    part2 = Part2();
                    break;
                default:
                    part1 = Part1();
                    part2 = Part2();
                    break;
            }

            var ms = Math.Round((DateTime.Now - start).TotalMilliseconds);

            Console.WriteLine($"Day {day} ({ms}ms):");
            if (part1 != "") Console.WriteLine($"    {part1}");
            if (part2 != "") Console.WriteLine($"    {part2}");
        }

        static string Part1()
        {
            long result = 0;

            var start = DateTime.Now;

            #region Solution
            var prev = 0;
            var curr = 0;
            for(int i = 0; i<inputs.Count; i++) {
                curr = inputs[i];
                if(i == 0)
                {

                }
                else
                {
                    if (curr > prev)
                    {
                        result++;
                    }
                }
                prev = curr;
            }
  
            #endregion

            var ms = Math.Round((DateTime.Now - start).TotalMilliseconds);

            if (result > 0) Answer1 = result.ToString();
            return $"Part 1 ({ms}ms): {result} ";
        }

        static string Part2()
        {
            long result = 0;

            var start = DateTime.Now;

            #region Solution
            var prev2 = 0;
            var prev = 0;
            var curr = 0;
            var sum = 999999999;
            var prevsum = 0;
            for (int i = 0; i < inputs.Count; i++)
            {
                curr = inputs[i];
                if (i == 0 ^ i == 1)
                {

                }
                else
                {
                    sum = prev2+ prev + curr;
                    if (sum > prevsum)
                    {
                        result++;
                    }
                    prevsum = sum;
                }
                prev2 = prev;
                prev = curr;
            }
            result = result - 1;

            #endregion

            var ms = Math.Round((DateTime.Now - start).TotalMilliseconds);

            if (result > 0) Answer2 = result.ToString();
            return $"Part 2 ({ms}ms): {result} ";
        }
    }
}