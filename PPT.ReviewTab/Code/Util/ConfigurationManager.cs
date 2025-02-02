using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using Newtonsoft.Json;
using PPT.ReviewTab.Code.Model;

namespace PPT.ReviewTab.Code.Util
{
    public sealed class ConfigurationManager
    {
        private const string ConfigFileName = "Configuration.json";
        private static readonly string AppDataFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), GlobalConst.BrandName, GlobalConst.AppName);

        private static readonly Lazy<ConfigurationManager> _instance = new Lazy<ConfigurationManager>(() => new ConfigurationManager());

        public static ConfigurationManager Instance => _instance.Value;

        public Configuration Configuration;

        public event EventHandler ConfigurationChanged;

        public ConfigurationManager() { 
            ConfigurationChanged += OnConfigurationChanged;
        }

        private void OnConfigurationChanged(object sender, EventArgs e)
        {
            SaveConfiguration();
        }

        public void NotifyConfigurationChanged(object sender)
        {
            ConfigurationChanged?.Invoke(sender, EventArgs.Empty);
        }



        /// <summary>
        /// Loads the configuration from the JSON file. Creates and saves a default configuration if the file does not exist.
        /// </summary>
        /// <returns>The loaded configuration.</returns>
        public void LoadConfiguration()
        {
            try
            {
                // Ensure the AppData folder exists
                if (!Directory.Exists(AppDataFolder))
                {
                    Directory.CreateDirectory(AppDataFolder);
                }

                // Full path to the configuration file
                string configFilePath = Path.Combine(AppDataFolder, ConfigFileName);

                // Check if the configuration file exists
                if (File.Exists(configFilePath))
                {
                    // Read and deserialize the configuration
                    string json = File.ReadAllText(configFilePath);
                    Configuration conf = JsonConvert.DeserializeObject<Configuration>(json);

                    if (conf.DefaultColorScheme == null)
                    {
                        conf.DefaultColorScheme = new ColorScheme(
                                    new Color(50, 255, 50),
                                    new Color(0, 0, 0),
                                    new Color(0, 0, 0));
                    }

                    Configuration = JsonConvert.DeserializeObject<Configuration>(json);

                    return;
                }
                else
                {
                    // Create default configuration
                    Configuration = CreateDefaultConfiguration();

                    // Save default configuration to the file
                    SaveConfiguration();

                    return;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading configuration: {ex.Message}");
                throw;
            }
        }

        


        /// <summary>
        /// Saves the configuration to the default file path.
        /// </summary>
        /// <param name="configuration">The configuration to save.</param>
        public void SaveConfiguration()
        {
            // Ensure the AppData folder exists
            if (!Directory.Exists(AppDataFolder))
            {
                Directory.CreateDirectory(AppDataFolder);
            }

            // Full path to the configuration file
            string filePath = Path.Combine(AppDataFolder, ConfigFileName);

            string json = JsonConvert.SerializeObject(Configuration, Newtonsoft.Json.Formatting.Indented);
            File.WriteAllText(filePath, json);
        }



        /// <summary>
        /// Creates a default configuration.
        /// </summary>
        /// <returns>The default configuration.</returns>
        private static Configuration CreateDefaultConfiguration()
        {
            int gap = 20;
            int width = 150;

            ColorScheme defaultColorScheme = new ColorScheme(new Color("#eee"), new Color("#000"), new Color("#000"));

            return new Configuration
            {
                DefaultColorScheme = defaultColorScheme,
                ItemGroups = new List<ItemGroup>
                {
                    new ItemGroup
                    {
                        Name = "Owner",
                        Items = new List<Item>
                        {
                            new Item { Name = "Name 1", ColorScheme=defaultColorScheme },
                            new Item { Name = "Name 2", ColorScheme=defaultColorScheme },
                            new Item { Name = "Name 3", ColorScheme=defaultColorScheme },
                            new Item { Name = "Name 4", ColorScheme=defaultColorScheme },
                            new Item { Name = "Name 5", ColorScheme=defaultColorScheme },
                            new Item { Name = "Name 6", ColorScheme=defaultColorScheme },
                            new Item { Name = "Name 7", ColorScheme=defaultColorScheme },
                            new Item { Name = "Name 8", ColorScheme=defaultColorScheme },
                            new Item { Name = "Name 9", ColorScheme=defaultColorScheme },
                            new Item { Name = "Name 10", ColorScheme=defaultColorScheme },
                            new Item { Name = "Name 11", ColorScheme=defaultColorScheme },
                            new Item { Name = "Name 12", ColorScheme=defaultColorScheme },
                        },
                        Shape = new GroupShape { 
                            Top = 10,
                            Right = 2*gap + width,
                            Width = width,
                            Hight = 30,
                            FontSize = 16,
                            ColorScheme = new ColorScheme( 
                                new Color(255, 255, 100),
                                new Color(0, 0, 0),
                                new Color(0,0,0))
                        }
                    },
                    new ItemGroup
                    {
                        Name = "Status",
                        Items = new List<Item>
                        {
                            new Item { Name = "Not started", ColorScheme = new ColorScheme { BackgroundColor = new Color("#f00"), TextColor = new Color("#fff"), FrameColor = new Color(0, 0, 0) } },
                            new Item { Name = "WIP", ColorScheme = new ColorScheme { BackgroundColor = new Color("#FFD961"), TextColor = new Color("#000"), FrameColor = new Color(0, 0, 0) } },
                            new Item { Name = "Beautify!", ColorScheme = new ColorScheme { BackgroundColor = new Color(224, 17, 224), TextColor = new Color("#fff"), FrameColor = new Color(0, 0, 0) } },
                            new Item { Name = "Draft", ColorScheme = new ColorScheme { BackgroundColor = new Color(9, 71, 110), TextColor = new Color("#fff"), FrameColor = new Color(0, 0, 0) } },
                            new Item { Name = "Done", ColorScheme = new ColorScheme { BackgroundColor = new Color("#00B050"), TextColor = new Color("#000"), FrameColor = new Color(0, 0, 0) } },
                            new Item { Name = "Final", ColorScheme = new ColorScheme { BackgroundColor = new Color(18, 153, 235), TextColor = new Color(0, 0, 0), FrameColor = new Color(0, 0, 0) } },
                            new Item { Name = "In Review", ColorScheme = new ColorScheme { BackgroundColor = new Color(200, 200, 200), TextColor = new Color(0, 0, 0), FrameColor = new Color(0, 0, 0) } },
                            new Item { Name = "Backup", ColorScheme = new ColorScheme { BackgroundColor = new Color(0, 0, 0), TextColor = new Color("#fff"), FrameColor = new Color(0, 0, 0) } },
                            new Item { Name = "N/A", ColorScheme=defaultColorScheme },
                        },
                        Shape = new GroupShape {
                            Top = 10,
                            Right = gap,
                            Width = width,
                            Hight = 30,
                            FontSize = 16,
                            ColorScheme = defaultColorScheme
                        }
                    }
                }
            };
        }
    }
}
