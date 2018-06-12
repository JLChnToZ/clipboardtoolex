using System;
using System.ServiceModel;

namespace ClipboardManager {
    internal static class SingleInstanceApplicationHandler {
        private const string Address = "net.pipe://localhost/IsClipboardExServiceAlive";
        private static ServiceHost host;

        [ServiceContract]
        private interface ITestLaunch {
            [OperationContract]
            bool Test();
        }

        private class TestLaunch: ITestLaunch {
            public bool Test() { return true; }
        }

        public static bool IsAppRunning() {
            ChannelFactory<ITestLaunch> channel = new ChannelFactory<ITestLaunch>(
                new NetNamedPipeBinding(),
                new EndpointAddress(Address)
            );
            ITestLaunch proxy = channel.CreateChannel();
            try {
                return proxy.Test();
            } catch (EndpointNotFoundException) {
                return false;
            }
        }

        public static void StartService() {
            if (host != null) return;
            host = new ServiceHost(typeof(TestLaunch), new Uri(Address));
            host.AddServiceEndpoint(typeof(ITestLaunch), new NetNamedPipeBinding(), "");
            host.Open();
        }

        public static void StopService() {
            if (host == null) return;
            host.Close();
            host = null;
        }
    }
}
