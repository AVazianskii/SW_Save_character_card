using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using OfficeOpenXml;
using SW_Character_creation;



namespace Character_design
{
    public class Save_character_excel
    {
        private static Save_character_excel Character_instance;



        private string player_cards_directory,
                       player_card_template,
                       character_directory,
                       character_file;



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
            character_file = character_directory + $"\\{Character.GetInstance().Name}" + ".xlsx";

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
                File.Copy(player_card_template, character_file, true);

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (var package = new ExcelPackage(new FileInfo(character_file)))
                {
                    byte row_index = 0;
                    byte skill_coloumn_num = 0;
                    byte skill_score_coloumn_num = 0;
                    // Заполняем поле общей информации о персонаже
                    ExcelWorksheet Character_card = package.Workbook.Worksheets[0];
                    Character_card.Cells[2, 1].Value = Character.GetInstance().Name + ", " + Character.GetInstance().Sex;
                    Character_card.Cells[3, 1].Value = Character.GetInstance().Character_race.Get_race_name();
                    Character_card.Cells[4, 2].Value = Character.GetInstance().Age.ToString();
                    Character_card.Cells[4, 4].Value = Character.GetInstance().Karma.ToString();
                    Character_card.Cells[5, 2].Value = Character.GetInstance().Range.Get_range_name();
                    Character_card.Cells[5, 4].Value = Character.GetInstance().Experience_left.ToString();
                    Character_card.Cells[16, 2].Value = Character.GetInstance().Name;

                    // Заполняем поля атрибутов
                    Character_card.Cells[09, 3].Value = Character.GetInstance().Strength.Get_atribute_score().ToString();
                    Character_card.Cells[10, 3].Value = Character.GetInstance().Stamina.Get_atribute_score().ToString();
                    Character_card.Cells[11, 3].Value = Character.GetInstance().Agility.Get_atribute_score().ToString();
                    Character_card.Cells[12, 3].Value = Character.GetInstance().Quickness.Get_atribute_score().ToString();
                    Character_card.Cells[13, 3].Value = Character.GetInstance().Intelligence.Get_atribute_score().ToString();
                    Character_card.Cells[14, 3].Value = Character.GetInstance().Perception.Get_atribute_score().ToString();
                    Character_card.Cells[15, 3].Value = Character.GetInstance().Charm.Get_atribute_score().ToString();
                    Character_card.Cells[16, 3].Value = Character.GetInstance().Willpower.Get_atribute_score().ToString();

                    // Заполняем поля боевых параметров
                    foreach (Skill_Class skill in Character.GetInstance().Skills)
                    {
                        switch(skill.ID)
                        {
                            case 16: Character_card.Cells[09, 5].Value = skill.Score; break;
                            case 17: Character_card.Cells[11, 5].Value = skill.Score; break;
                            case 12: Character_card.Cells[13, 5].Value = skill.Score; break;
                            case 13: Character_card.Cells[15, 5].Value = skill.Score; break;
                            case 10: Character_card.Cells[17, 5].Value = skill.Score; break;
                            case 09: Character_card.Cells[09, 7].Value = skill.Score; break;
                            case 11: Character_card.Cells[11, 7].Value = skill.Score; break;
                            case 07: Character_card.Cells[13, 7].Value = skill.Score; break;
                            case 18: Character_card.Cells[15, 7].Value = skill.Score; break;
                            case 15: Character_card.Cells[17, 7].Value = skill.Score; break;
                        }
                    }

                    // Заполняем поля пирамиды ранений и штрафов за них
                    Character_card.Cells[09, 08].Value = Character.GetInstance().Scratch_lvl.ToString();
                    Character_card.Cells[19, 08].Value = Character.GetInstance().Scratch_penalty.ToString();
                    Character_card.Cells[10, 09].Value = Character.GetInstance().Light_wound_lvl.ToString();
                    Character_card.Cells[18, 09].Value = Character.GetInstance().Light_wound_penalty.ToString();
                    Character_card.Cells[11, 10].Value = Character.GetInstance().Medium_wound_lvl.ToString();
                    Character_card.Cells[17, 10].Value = Character.GetInstance().Medium_wound_penalty.ToString();
                    Character_card.Cells[12, 11].Value = Character.GetInstance().Tough_wound_lvl.ToString();
                    Character_card.Cells[16, 11].Value = Character.GetInstance().Tough_wound_penalty.ToString();
                    Character_card.Cells[14, 12].Value = Character.GetInstance().Scratch_lvl.ToString();

                    // Заполняем расчитанные боевые параметры
                    Character_card.Cells[11, 14].Value = Character.GetInstance().Reaction.ToString();
                    Character_card.Cells[13, 14].Value = Character.GetInstance().Armor.ToString();
                    Character_card.Cells[15, 14].Value = Character.GetInstance().Force_resistance.ToString();
                    Character_card.Cells[17, 14].Value = Character.GetInstance().Hideness.ToString();
                    Character_card.Cells[19, 14].Value = Character.GetInstance().Watchfullness.ToString();
                    Character_card.Cells[21, 14].Value = Character.GetInstance().Concentration.ToString();

                    // Заполняем боевые формы
                    
                    if (Character.GetInstance().Combat_sequences_with_points.Count > 0)
                    {
                        row_index = 36;
                        foreach (Abilities_sequence_template sequence in Character.GetInstance().Combat_sequences_with_points)
                        {
                            Character_card.Cells[row_index, 08].Value = sequence.Name;
                            Character_card.Cells[row_index, 13].Value = sequence.Level;
                            row_index = (byte)(row_index + 1);
                        }
                    }

                    // Заполняем формы Силы
                    if (Character.GetInstance().Forceuser)
                    {
                        if (Character.GetInstance().Force_sequences_with_points.Count > 0)
                        {
                            row_index = 36;
                            foreach (Abilities_sequence_template sequence in Character.GetInstance().Force_sequences_with_points)
                            {
                                Character_card.Cells[row_index, 08].Value = sequence.Name;
                                Character_card.Cells[row_index, 13].Value = sequence.Level;
                                row_index = (byte)(row_index + 1);
                            }
                        }
                    }

                    // Заполняем поля навыков
                   
                    if (Character.GetInstance().Skills_with_points.Count > 0)
                    {
                        row_index = 6;
                        foreach (Skill_Class skill in Character.GetInstance().Skills_with_points)
                        {
                            if (row_index < 21)
                            {
                                skill_coloumn_num = 18;
                                skill_score_coloumn_num = 21;
                            }
                            else
                            {
                                row_index = 6;
                                skill_coloumn_num = 22;
                                skill_score_coloumn_num = 25;
                            }
                            Character_card.Cells[row_index, skill_coloumn_num].Value = skill.Name;
                            Character_card.Cells[row_index, skill_score_coloumn_num].Value = skill.Score;
                            row_index = (byte)(row_index + 1);
                        }
                    }

                    // Заполняем поля навыков Силы
                    if (Character.GetInstance().Force_skills_with_points.Count > 0)
                    {
                        skill_coloumn_num = 16;
                        skill_score_coloumn_num = 17;
                        row_index = 6;
                        foreach (Force_skill_class skill in Character.GetInstance().Force_skills_with_points)
                        {
                            Character_card.Cells[row_index, skill_coloumn_num].Value = skill.Name;
                            Character_card.Cells[row_index, skill_score_coloumn_num].Value = skill.Score;
                            row_index = (byte)(row_index + 1);
                        }
                    }

                    // Заполняем положительные особенности
                    if (Character.GetInstance().Positive_features_with_points.Count > 0)
                    {
                        skill_coloumn_num = 16;
                        row_index = 23;
                        foreach (All_feature_template feature in Character.GetInstance().Positive_features_with_points)
                        {
                            Character_card.Cells[row_index, skill_coloumn_num].Value = feature.Name;
                            row_index = (byte)(row_index + 1);
                        }
                    }

                    // Заполняем отрицательные особенности
                    if (Character.GetInstance().Negative_features_with_points.Count > 0)
                    {
                        skill_coloumn_num = 21;
                        row_index = 23;
                        foreach (All_feature_template feature in Character.GetInstance().Negative_features_with_points)
                        {
                            Character_card.Cells[row_index, skill_coloumn_num].Value = feature.Name;
                            row_index = (byte)(row_index + 1);
                        }
                    }

                    package.Save();
                    //package.SaveAs();
                }
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



        private void CheckExistCharacterCard()
        {

        }
    }
}
