using System;
using System.IO;
using Xbim.Ifc;
using Xbim.Ifc4.Interfaces;
using System.Linq;
using System.Collections.Generic;
using Xbim.Ifc2x3.GeometryResource;
using Xbim.Ifc.Extensions;
using Newtonsoft.Json;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Web;

class Program
{
    [STAThread]
    static void Main(string[] args)
    {

        string connectionString = @"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=|DataDirectory|\ElementDB.mdf;Integrated Security=True";

        using (SqlConnection connection = new SqlConnection(connectionString))
        {
            connection.Open();

            // List of variables looked for in the ifc model
            // Units unclear!
            var wantedProperties = new List<string>
            {
                "Weight",
                "Height",
                "Length",
                "Width",
                "Volume",
                "Gross footprint area",
                "Area per tons",
                "Net surface area",
                "Bottom elevation",
                "Top elevation"
            };

            // Ifc file selection window popup
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Title = "Select a ifc model";
                openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + "\\Downloads";
                openFileDialog.Filter = "Ifc Files (*.ifc)|*.ifc";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    using (var model = IfcStore.Open(openFileDialog.FileName))
                    {
                        // Check if project exists
                        string projectName = Path.GetFileNameWithoutExtension(openFileDialog.FileName);
                        string checkProject = "SELECT * FROM Projects WHERE ProjectName = '" + projectName + "';";
                        SqlCommand checkCommand = new SqlCommand(checkProject, connection);
                        int result = checkCommand.ExecuteNonQuery();
                        
                        if (result == -1) 
                        {
                            Console.WriteLine("Project not in database, adding it.");
                            SqlCommand addCommand = new SqlCommand("INSERT INTO Projects (ProjectName) VALUES ('" + projectName + "')", connection);
                            int addResult = addCommand.ExecuteNonQuery();
                            if (addResult == -1)
                            {
                                Console.WriteLine("Inserting new project failed");
                            }
                        }


                        // Create list of all elements (Empty for now)
                        var allElements = new List<IIfcElement>();

                        // Fill the list of allElements with all of the selected elements (Reduces the amount of loops needed)
                        allElements.AddRange(model.Instances.OfType<IIfcBeam>().ToList());
                        allElements.AddRange(model.Instances.OfType<IIfcColumn>().ToList());
                        allElements.AddRange(model.Instances.OfType<IIfcSlab>().ToList());

                        // Loop through the elements
                        foreach (var element in allElements)
                        {
                            // Get the element tye name
                            var elementTypeName = element.GetType().Name;

                            string addElement = $"INSERT INTO {elementTypeName}";

                            // Print the basic information
                            Console.WriteLine($"{elementTypeName} ID: {element.GlobalId}");
                            Console.WriteLine($"{elementTypeName} name: {element.Name}");

                            Dictionary<string, string> elementProperties = new Dictionary<string, string>();

                            string columnNames = "(" +
                                "uniqueId, " +
                                "Name, " +
                                "Weight, " +
                                "Volume, " +
                                "GrossFootprintArea, " +
                                "AreaPerTons, " +
                                "NetSurfaceArea, " +
                                "Height, " +
                                "Width, " +
                                "Length, " +
                                "BottomElevation, " +
                                "TopElevation, " +
                                "ProjectName)";

                            string columnValues = $"('{element.GlobalId}', '{element.Name}', ";

                            // Get all single-value properties of the beam
                            var properties = element.IsDefinedBy
                                .Where(r => r.RelatingPropertyDefinition is IIfcPropertySet)
                                .SelectMany(r => ((IIfcPropertySet)r.RelatingPropertyDefinition).HasProperties)
                                .OfType<IIfcPropertySingleValue>();

                            foreach (var property in properties)
                            {
                                if (wantedProperties.Contains(property.Name))
                                {
                                    columnValues += $"'{property.NominalValue}', ";
                                    Console.WriteLine($"{elementTypeName} {property.Name}: {property.NominalValue}");
                                }
                            }
                            columnValues += $"'{projectName}')";
                            // Get element location
                            var location = element.ObjectPlacement.ToMatrix3D().Translation;
                            Console.WriteLine($"{elementTypeName} {element.Name} is located at {location}");

                            Console.WriteLine();
                            SqlCommand addElementCommand = new SqlCommand("INSERT INTO " + elementTypeName + "s " + columnNames + " VALUES " + columnValues + "", connection);
                            int resultAddElements = addElementCommand.ExecuteNonQuery();

                        }
                    }
                }

                connection.Close();
                Console.WriteLine("Press any key to exit...");
                Console.ReadKey();
            }
        }
    }
}