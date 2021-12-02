using AdventOfCode.Helpers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AdventOfCode.Y2021.Days
{
    public static class Day02
    {
        static int day = 2;
        static List<string>? inputs;

        public static string? Answer1 { get; set; }
        public static string? Answer2 { get; set; }

        public static void Run(int part, bool test)
        {
            inputs = InputManager.GetInputAsStrings(day, test);

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
            int depth = 0;
            int forward = 0;

            for (int i = 0; i < inputs.Count; i++)
            {
                string s = inputs[i];
                string[] ss = s.Split();
                string command = ss[0];
                //int amt = (int)ss[1];
                int amt = int.Parse(ss[1]);

                if (command == "forward") forward += amt;
                if (command == "up") depth -= amt;
                if (command == "down") depth += amt;
            }
            result = depth * forward;
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
            int depth = 0;
            int forward = 0;
            int aim = 0;

            for (int i = 0; i < inputs.Count; i++)
            {
                string s = inputs[i];
                string[] ss = s.Split();
                string command = ss[0];
                //int amt = (int)ss[1];
                int amt = int.Parse(ss[1]);

                if (command == "forward") { forward += amt; depth += aim * amt; }
                else if (command == "up") aim -= amt;
                else if (command == "down") aim += amt;
            }
            result = depth * forward;


            #endregion

            var ms = Math.Round((DateTime.Now - start).TotalMilliseconds);

            if (result > 0) Answer2 = result.ToString();
            return $"Part 2 ({ms}ms): {result} ";
        }
    }
}