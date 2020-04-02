using System;
using System.Diagnostics;

namespace Microsoft.Crm.Services.Utility
{
	internal sealed class CommandLineParameters
	{
		private const char StartCharacter = '/';
		private const char SeparatorCharacter = ':';

		#region Properties
		public OperationType Operation { get; private set; }

		public string DeviceName { get; private set; }

		public string DevicePassword { get; private set; }

		public bool ForceRegistration { get; private set; }

		public Uri IssuerUri { get; private set; }
		#endregion

		#region Methods
		public bool Parse(string[] args)
		{
			if (null == args)
			{
				throw new ArgumentNullException("args");
			}

			bool isValid = true;
			bool isOperationValid = false;
			foreach (string arg in args)
			{
				if (string.IsNullOrWhiteSpace(arg))
				{
					continue;
				}

				if (StartCharacter == arg[0])
				{
					int separatorPosition = arg.IndexOf(SeparatorCharacter, 1);
					if (-1 != separatorPosition)
					{
						string name = arg.Substring(1, separatorPosition - 1);
						string value = arg.Substring(separatorPosition + 1);
						if (!string.IsNullOrWhiteSpace(value))
						{
							switch (name.ToUpperInvariant())
							{
								case "OPERATION":
								case "O":
									{
										OperationType result;
										if (Enum.TryParse<OperationType>(value, true, out result))
										{
											this.Operation = result;
											isOperationValid = true;
											continue;
										}
									}
									break;
								case "NAME":
								case "N":
									this.DeviceName = value;
									continue;
								case "PASSWORD":
								case "P":
									this.DevicePassword = value;
									continue;
								case "ISSUER":
								case "URL":
									try
									{
										this.IssuerUri = new Uri(value);
									}
									catch (UriFormatException)
									{
										//Ignore this exception
									}
									continue;
								case "FORCE":
								case "F":
									this.ForceRegistration = true;
									continue;
							}
						}
					}
				}

				Console.Error.WriteLine("Invalid Argument: \"{0}\"", arg);
				isValid = false;
			}

			if (isValid && (!isOperationValid || string.IsNullOrWhiteSpace(this.DeviceName) != string.IsNullOrWhiteSpace(this.DevicePassword)))
			{
				isValid = false;
			}

			return isValid;
		}

		public void ShowHelp()
		{
			Console.Out.WriteLine("{0}", Process.GetCurrentProcess().ProcessName);
			Console.Out.WriteLine(" /operation:<operation> - Valid Options are Register or Show. Required.");
			Console.Out.WriteLine(" /name:<device name> - Optional.");
			Console.Out.WriteLine(" /password:<device password> - Optional.");
		}
		#endregion
	}

	internal enum OperationType
	{
		Register,
		Show
	}
}
