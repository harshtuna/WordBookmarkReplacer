using System;
using System.Collections.Generic;
using System.CommandLine;
using System.CommandLine.Invocation;
using System.IO;

namespace MsWordTemplateProcessor.Cli
{
    public static class Program
    {
        // see https://github.com/dotnet/command-line-api/wiki
        /// <summary>
        /// Replace MS Word bookmarks with text
        /// </summary>
        public static int Main(string[] args)
        {
            // Create a root command with some options
            var rootCommand = new RootCommand
            {
                new Option<FileInfo>(new[] {"--template", "-t"}, "Template file")
                {
                    Argument = new Argument<FileInfo>().ExistingOnly()
                },
                new Option<FileInfo>(new[] {"--output", "-o"}, "Output file"),
                new Option<List<string>>(new[] {"--bookmark", "-b"}, () => new List<string>(), "Bookmark to replace"),
                new Option<List<string>>(new[] {"--value", "-v"}, () => new List<string>(), "Value to insert"),
            };

            rootCommand.Description = "Replace MS Word bookmarks with text";

            // Note that the parameters of the handler method are matched according to the names of the options
            rootCommand.Handler = CommandHandler.Create<FileInfo, FileInfo, string[], string[]>(
                (template, output, bookmark, value) =>
                    ProcessTemplate(template.FullName, output.FullName, bookmark, value));

            // Parse the incoming args and invoke the handler
            return rootCommand.InvokeAsync(args).Result;
        }

        private static void ProcessTemplate(string template, string output, string[] bookmark, string[] value)
        {
            var templateProcessor = new BookmarkTemplateProcessor(template, output);
            for (var i = 0; i < bookmark.Length; i++) templateProcessor.ApplyValue(bookmark[i], value[i]);
            templateProcessor.Document.Save();
            templateProcessor.Document.Close();
        }
    }
}