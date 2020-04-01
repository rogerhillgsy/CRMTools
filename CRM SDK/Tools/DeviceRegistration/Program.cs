using System;
using System.ServiceModel.Description;

namespace Microsoft.Crm.Services.Utility
{
	internal sealed class LiveDeviceIdUtil
	{
		public static void Main(string[] args)
		{
			CommandLineParameters parameters = new CommandLineParameters();
			if (!parameters.Parse(args))
			{
				parameters.ShowHelp();
				return;
			}

			ClientCredentials credentials;
			switch (parameters.Operation)
			{
				case OperationType.Register:
					credentials = DeviceIdManager.LoadDeviceCredentials(parameters.IssuerUri);
					if (null != credentials && !parameters.ForceRegistration)
					{
						Console.Error.WriteLine("Error: Device is already registered.");
						break;
					}

					try
					{
						credentials = DeviceIdManager.RegisterDevice(Guid.NewGuid(), parameters.IssuerUri,
							parameters.DeviceName, parameters.DevicePassword);
					}
					catch (DeviceRegistrationFailedException ex)
					{
						ConsoleColor originalColor = Console.ForegroundColor;
						try
						{
							Console.ForegroundColor = ConsoleColor.Red;
							Console.Error.WriteLine("Error - {1}", ex.RegistrationErrorCode, ex.Message);
						}
						finally
						{
							Console.ForegroundColor = originalColor;
						}
					}
					break;
				case OperationType.Show:
					credentials = DeviceIdManager.LoadDeviceCredentials(parameters.IssuerUri);
					break;
				default:
					throw new NotImplementedException("Operation = " + parameters.Operation);
			}

			if (null == credentials)
			{
				Console.Out.WriteLine("Device is not registered.");
			}
			else
			{
				Console.Out.WriteLine("Device ID: {0}", credentials.UserName.UserName);
				Console.Out.WriteLine("Device Password: {0}", credentials.UserName.Password);
			}
		}
	}
}
