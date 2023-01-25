using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Windows;



namespace Character_design
{
    public class Save_character_excel
    {
        private static Save_character_excel Character_instance;



        private string player_cards_directory,
                       player_card_template,
                       character_directory;



        public static Save_character_excel GetInstance()
        {
            if (Character_instance == null)
            {
                Character_instance = new Save_character_excel();
            }
            return Character_instance;
        }
        public void Save_character_to_Excel_card (out string error_msg)
        {
            error_msg = "";
            character_directory = player_cards_directory + $"\\{Character.GetInstance().Name}";


            if (File.Exists(player_card_template))
            {
                // директория с карточками уже созданных персонажей
                if (Directory.Exists(player_cards_directory) == false)
                {
                    Directory.CreateDirectory(player_cards_directory);
                }
                // директория с карточкой конкретного персонажа
                if (Directory.Exists(character_directory) == false)
                {
                    Directory.CreateDirectory(character_directory);
                }
                File.Copy(player_card_template, character_directory + $"\\{Character.GetInstance().Name}" + ".xlsx");
            }
            else
            {
                error_msg = "Создание анкеты персонажа невозможно! Отсутствует шаблон анкеты.";
            }
        }



        private Save_character_excel()
        {
            player_cards_directory = Directory.GetCurrentDirectory() + "\\Player_cards";
            player_card_template = Directory.GetCurrentDirectory() + "\\Player_card_template\\Template_v2.xlsx";
        }
    }
}
