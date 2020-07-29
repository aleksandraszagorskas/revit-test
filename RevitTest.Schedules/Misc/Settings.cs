using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RevitTest.Schedules.Misc
{
    public class Settings
    {
        public string ROOT_FOLDER { get; private set; } = null;
        public string IMAGES_FOLDER { get; private set; } = null;
        public string EXPORT_FOLDER { get; private set; } = null;
        public string EXCEL_EXPORT_FOLDER { get; private set; } = null;
        public string EXCEL_TEMPLATE_PATH { get; private set; } = null;
        public string EXCEL_TEST_DATA_PATH { get; private set; } = null;
        public string EXCEL_TEST_DATA_2_PATH { get; private set; } = null;

        public float IMAGE_ROW_HEIGHT { get; private set; } = -1.0f;
        public float IMAGE_COLUMN_WIDTH { get; private set; } = -1.0f;
        public float IMAGE_PADDING { get; private set; } = -1.0f;

        public bool GENERATE_IMAGES { get; private set; } = false;

        public int PAGE_ITEM_COUNT { get; private set; } = -1;

        public string SHARED_PARAMETERS_FOLDER { get; private set; } = null;
        public string SHARED_PARAMETER_FILE_NAME { get; private set; } = null;
        public string SHARED_PARAMETER_GROUP_NAME { get; private set; } = null;
        public string TITLEBLOCK_FOLDER { get; private set; } = null;
        public string MAIN_TITLEBLOCK_NAME { get; private set; } = null;
        public string MAIN_TITLEBLOCK_FILE_NAME { get; private set; } = null;
        public string MAIN_LIST_NAME { get; private set; } = null;
        public string EXPANDABLE_TITLEBLOCK_NAME { get; private set; } = null;
        public string EXPANDABLE_TITLEBLOCK_FILE_NAME { get; private set; } = null;
        public string EXPANDABLE_LIST_NAME { get; private set; } = null;
        public string ANNOTATIONS_FOLDER { get; private set; } = null;
        public string DEFAULT_ANNOTATION_FILE_NAME { get; private set; } = null;
        public string DEFAULT_ANNOTATION_IMAGE_FILE_NAME { get; private set; } = null;
        public string TEMPLATES_FOLDER { get; private set; } = null;
        public string SCHEDULE_TEMPLATE_FILE_NAME { get; private set; } = null;

        #region Singleton
        private Settings()
        {
            ROOT_FOLDER = @"C:\Users\AleksandrasZagorskas\Desktop\revit_csd\schedules";
            IMAGES_FOLDER = Path.Combine(ROOT_FOLDER, "images");
            EXPORT_FOLDER = Path.Combine(ROOT_FOLDER, "export");
            EXCEL_EXPORT_FOLDER = Path.Combine(EXPORT_FOLDER, "excel");
            EXCEL_TEMPLATE_PATH = Path.Combine(EXCEL_EXPORT_FOLDER, "template3.xlsx");
            EXCEL_TEST_DATA_PATH = Path.Combine(EXCEL_EXPORT_FOLDER, "test-data.csv");
            EXCEL_TEST_DATA_2_PATH = Path.Combine(EXCEL_EXPORT_FOLDER, "test-data2.csv");

            IMAGE_ROW_HEIGHT = 200.0f;
            IMAGE_COLUMN_WIDTH = 50.0f;
            IMAGE_PADDING = 1.0f;

            GENERATE_IMAGES = false;

            PAGE_ITEM_COUNT = 12;

            SHARED_PARAMETERS_FOLDER = Path.Combine(ROOT_FOLDER, "shared-params");
            SHARED_PARAMETER_FILE_NAME = Path.Combine(SHARED_PARAMETERS_FOLDER, "HelloSharedParameterWorld.txt");
            SHARED_PARAMETER_GROUP_NAME = "My Shared Parameter Group";
            TITLEBLOCK_FOLDER = Path.Combine(ROOT_FOLDER, "titleblocks");
            MAIN_TITLEBLOCK_NAME = "Test Titile Block";
            MAIN_TITLEBLOCK_FILE_NAME = Path.Combine(TITLEBLOCK_FOLDER, $"{MAIN_TITLEBLOCK_NAME}.rfa");
            MAIN_LIST_NAME = "Title schedule";
            EXPANDABLE_TITLEBLOCK_NAME = "Test Expandable Titile Block";
            EXPANDABLE_TITLEBLOCK_FILE_NAME = Path.Combine(TITLEBLOCK_FOLDER, $"{EXPANDABLE_TITLEBLOCK_NAME}.rfa");
            EXPANDABLE_LIST_NAME = "Expandable schedule";
            ANNOTATIONS_FOLDER = Path.Combine(ROOT_FOLDER, "annotations");
            DEFAULT_ANNOTATION_FILE_NAME = Path.Combine(ANNOTATIONS_FOLDER, "rebar-test-image.rfa");
            DEFAULT_ANNOTATION_IMAGE_FILE_NAME = Path.Combine(IMAGES_FOLDER, $"{Path.GetFileNameWithoutExtension(DEFAULT_ANNOTATION_FILE_NAME)}.png");
            TEMPLATES_FOLDER = Path.Combine(ROOT_FOLDER, "templates");
            SCHEDULE_TEMPLATE_FILE_NAME = Path.Combine(TEMPLATES_FOLDER, "REBAR_SCHEDULE_TEMPLATES.rvt");
        }

        private static Settings instance = null;
        public static Settings Instance
        {
            get
            {
                if (instance == null)
                {
                    return new Settings();
                }
                else
                {
                    return instance;
                }
            }
        }
        #endregion
    }
}
