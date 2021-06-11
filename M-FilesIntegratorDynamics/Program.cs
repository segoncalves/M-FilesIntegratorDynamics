using Integrador_MFD.Dynamics;
using System;

namespace DynamicsIntegrator
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");

            Connect connectDynamics = new Connect();
            connectDynamics.Process();
        }
    }
}
