using System;
using System.Collections.Generic;


namespace Character_design
{
    public class Save_character_excel
    {
        private static Save_character_excel Character_instance;



        public static Save_character_excel GetInstance()
        {
            if (Character_instance == null)
            {
                Character_instance = new Save_character_excel();
            }
            return Character_instance;
        }
        public void Save_character_to_Excel_card (Character character)
        {

        }


        private Save_character_excel()
        {

        }
    }
}
