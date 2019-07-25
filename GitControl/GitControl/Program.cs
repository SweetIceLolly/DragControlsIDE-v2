using Octokit;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Sockets;
using System.Text;
using System.Threading.Tasks;

namespace GitControl
{
    class Program
    {
        /*
         * args[1] = git bash path
         * args[2] = operation folder
         * args[3] = mode
         *           {
         *              auto                - let gitcontrol handle it
         *              custom              - take care of it yourself with git command
         *              
         * args[4] = command
         *           {
         *              init                - initalize the folder and create repo in github
         *              commit              - commit changes
         *              push                - push repo to github
         *              pull                - fetch changes from github and merge
         *              clone               - clone repo from github
         *           }
         * 
         * args[5] = username
         * args[6] = password
         * args[7] = appendix message1      - description of repo when args[4] is "init" or commit message when args[4] is "commit"
         * args[8] = appendix message2      - decide whether the repo is private or not when args[4] is "init"
         */
        static void Main(string[] args)
        {
            //connect github
            //GitHubClient client = new GitHubClient(new ProductHeaderValue("gitcontrol"));
            //client.Credentials = new Credentials(args[5], args[6]);

            //connect drag-controls IDE
            TcpClient sock = new TcpClient("127.0.0.1",6028);
            BinaryWriter writer = new BinaryWriter(sock.GetStream());
            writer.Write(" hello");

            sock.Close();
            //Process process = new Process();
            //string command = "";
            //string[] tem = args[2].Split('\\');
            //if(args[3]=="auto")
            //{
            //    switch (args[4])
            //    {
            //        case "init":
            //            command = "cd " + args[2] + "&&git init";
            //            NewRepository repo = new NewRepository(tem[tem.Length-1]);
            //            repo.Description = args[7];
            //            repo.Private = args[8] == "private" ? true : false;
            //            break;
            //        case "commit":
            //            command = "cd " + args[2] + "&&git add .&&git commit -m '" + args[7] + "'";
            //            break;
            //        case "push":
            //            command = "cd " + args[2] + "&&git push";
            //            break;
            //        case "pull":
            //            command = "cd " + args[2] + "&&git pull";
            //            break;
            //        case "clone":

            //            break;
            //    }
            //}
            //else
            //{
            //    command = args[4];
            //}
            //process.StartInfo = new ProcessStartInfo()
            //{
            //    FileName = args[1],
            //    UseShellExecute = false,
            //    RedirectStandardOutput = true,
            //    Arguments = command;
            //};

        }
    }
}
