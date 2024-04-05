using System;
using System.Linq;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using System.Xml.Serialization;
using System.Threading.Tasks;
using OfficeOpenXml;
using Spire.Xls;
using SW_Character_creation;



namespace Character_design
{
    public class Character_card
    {
        private string player_cards_directory,
                       player_card_template,
                       character_directory,
                       character_file,
                       test;
        private Character _character;


        public string Test { get; private set; }
        

        public void Edit_character_card_from_Excel(string character_card_path, Character character, Main_model model)
        {
            this._character = character;
            using (var package = new ExcelPackage(new FileInfo(character_card_path)))
            {
                ExcelWorksheet Character_card = package.Workbook.Worksheets[0];

                // Восстанавливаем имя персонажа
                _character.Name = Character_card.Cells[2, 17].Value.ToString();

                // Восстанавливаем расу персонажа
                var character_race = from race in model.Race_Manager.Get_Race_list()
                                     where race.Race_name == Character_card.Cells[3, 1].Value.ToString()
                                     select race;

                _character.Character_race = character_race.First();

                // Восстанавливаем ранг персонажа
                var character_ranges = from range in model.Range_Manager.Ranges()
                                        where range.Range_name == Character_card.Cells[5, 2].Value.ToString()
                                        select range;
                test = Character_card.Cells[5, 2].Value.ToString();

                character.Range = character_ranges.First();

                // Восстанавливаем возраст персонажа
                _character.Age = Convert.ToInt32(Character_card.Cells[4, 2].Value);
                /*  */

                // Восстанавливаем карму персонажа
                _character.Karma = Convert.ToInt32(Character_card.Cells[4, 4].Value);
                /* Надо проверить на принадлежность к сторонам силы, если адепт Силы */

                // Восстанавливаем остаток очков опыта персонажа
                _character.Experience_left = Convert.ToInt32(Character_card.Cells[5, 4].Value);

                /* Надо в шаблоне добавить остаток по очкам атрибутов */

                // Восстанавливаем атрибуты персонажа
                _character.Strength.Set_atr_score(Convert.ToInt32(Character_card.Cells[9, 3].Value));

                //Character_card.Cells[2, 1].Value = _character.Name + ", " + _character.Sex;
                //Character_card.Cells[3, 1].Value = _character.Character_race.Get_race_name();
                //Character_card.Cells[4, 2].Value = _character.Age.ToString();
                //Character_card.Cells[4, 4].Value = _character.Karma.ToString();
                //Character_card.Cells[5, 2].Value = _character.Range.Get_range_name();
                //Character_card.Cells[5, 4].Value = _character.Experience_left.ToString();
            }
        }
        public async Task Save_character_xmlAsync(Character character)
        {
            await Task.Run(() => Save_character_xml(character));
        }
        public async Task Save_character_to_Excel_cardAsync(Character character)
        {
            await Task.Run(() => Save_character_to_Excel_card(character));
        }



        public Character_card()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            player_cards_directory = Directory.GetCurrentDirectory() + "\\Карточки персонажей";
            player_card_template = Directory.GetCurrentDirectory() + "\\Player_card_template\\Template_v3.xlsx";
        }



        // временно сделан публичным
        public void Save_character_xml(Character character)
        {
            _character = character;
            XmlSerializer xml = new XmlSerializer(typeof(Character));

            using (FileStream fs = new FileStream(player_cards_directory + $"\\{_character.Name}\\{_character.Name}.xml", FileMode.OpenOrCreate))
            {
                xml.Serialize(fs, _character);
            }            
        }
        // временно сделан публичным
        public void Save_character_to_Excel_card(Character character)
        {
            this._character = character;
            character_directory = player_cards_directory + $"\\{_character.Name}";
            character_file = character_directory + $"\\{_character.Name}" + ".xlsx";

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
                // Копируем картинку персонажа в его папку
                string extension = "";
                if (_character.Img_path.Contains(".png"))
                {
                    extension = ".png";
                }
                else if (_character.Img_path.Contains(".jpg"))
                {
                    extension = ".jpg";
                }
                File.Copy(_character.Img_path, character_directory + $"\\{_character.Name}" + extension, true);

                using (var package = new ExcelPackage(new FileInfo(character_file)))
                {
                    byte row_index = 0;
                    byte skill_coloumn_num = 0;
                    byte skill_score_coloumn_num = 0;
                    // Заполняем поле общей информации о персонаже
                    ExcelWorksheet Character_card = package.Workbook.Worksheets[0];
                    Character_card.Cells[2, 1].Value = _character.Name + ", " + _character.Sex;
                    Character_card.Cells[3, 1].Value = _character.Character_race.Get_race_name();
                    Character_card.Cells[4, 2].Value = _character.Age.ToString();
                    Character_card.Cells[4, 4].Value = _character.Karma.ToString();
                    Character_card.Cells[5, 2].Value = _character.Range.Get_range_name();
                    Character_card.Cells[5, 4].Value = _character.Experience_left.ToString();
                    Character_card.Cells[2, 17].Value = _character.Name;

                    // Заполняем поля атрибутов
                    Character_card.Cells[09, 3].Value = _character.Strength.Get_atribute_score().ToString();
                    Character_card.Cells[11, 3].Value = _character.Stamina.Get_atribute_score().ToString();
                    Character_card.Cells[13, 3].Value = _character.Agility.Get_atribute_score().ToString();
                    Character_card.Cells[15, 3].Value = _character.Quickness.Get_atribute_score().ToString();
                    Character_card.Cells[17, 3].Value = _character.Intelligence.Get_atribute_score().ToString();
                    Character_card.Cells[19, 3].Value = _character.Perception.Get_atribute_score().ToString();
                    Character_card.Cells[21, 3].Value = _character.Charm.Get_atribute_score().ToString();
                    Character_card.Cells[23, 3].Value = _character.Willpower.Get_atribute_score().ToString();

                    // Загружаем картинку персонажа
                    var Character_picture = Character_card.Drawings.AddPicture("Character_picture", _character.Img_path);
                    Character_picture.SetPosition(0, 0, 12, 0);
                    // Конвертируем размер ячеек Экселя из мм в пиксели (1 мм = 4 пикселя)
                    Character_picture.SetSize(Convert.ToInt32(52 * 4), Convert.ToInt32(65 * 4));

                    // Заполняем поля боевых параметров
                    foreach (Skill_Class skill in _character.Skills)
                    {
                        switch (skill.ID)
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
                    Character_card.Cells[10, 09].Value = _character.Scratch_lvl.ToString();
                    Character_card.Cells[19, 09].Value = _character.Scratch_penalty.ToString();
                    Character_card.Cells[11, 10].Value = _character.Light_wound_lvl.ToString();
                    Character_card.Cells[18, 10].Value = _character.Light_wound_penalty.ToString();
                    Character_card.Cells[12, 11].Value = _character.Medium_wound_lvl.ToString();
                    Character_card.Cells[17, 11].Value = _character.Medium_wound_penalty.ToString();
                    Character_card.Cells[13, 12].Value = _character.Tough_wound_lvl.ToString();
                    Character_card.Cells[16, 12].Value = _character.Tough_wound_penalty.ToString();
                    Character_card.Cells[14, 13].Value = _character.Mortal_wound_lvl.ToString();

                    // Заполняем расчитанные боевые параметры
                    Character_card.Cells[11, 15].Value = _character.Reaction.ToString();
                    Character_card.Cells[13, 15].Value = _character.Armor.ToString();
                    Character_card.Cells[15, 15].Value = _character.Force_resistance.ToString();
                    Character_card.Cells[17, 15].Value = _character.Hideness.ToString();
                    Character_card.Cells[19, 15].Value = _character.Watchfullness.ToString();
                    if (_character.Forceuser)
                    {
                        Character_card.Cells[21, 15].Value = _character.Concentration.ToString();
                    }
                    else
                    {
                        Character_card.Cells[21, 15].Value = 0;
                    }

                    // Заполняем боевые формы
                    row_index = 41;
                    if (_character.Combat_sequences_with_points.Count > 0)
                    {

                        foreach (Abilities_sequence_template sequence in _character.Combat_sequences_with_points)
                        {
                            Character_card.Cells[row_index, 08].Value = sequence.Name;
                            Character_card.Cells[row_index, 14].Value = sequence.Level;
                            row_index = (byte)(row_index + 1);
                        }
                    }

                    // Заполняем формы Силы
                    if (_character.Forceuser)
                    {
                        if (_character.Force_sequences_with_points.Count > 0)
                        {
                            foreach (Abilities_sequence_template sequence in _character.Force_sequences_with_points)
                            {
                                Character_card.Cells[row_index, 08].Value = sequence.Name;
                                Character_card.Cells[row_index, 14].Value = sequence.Level;
                                row_index = (byte)(row_index + 1);
                            }
                        }
                    }

                    // Заполняем поля навыков

                    if (_character.Skills_with_points.Count > 0)
                    {
                        row_index = 5;
                        foreach (Skill_Class skill in _character.Skills_with_points)
                        {
                            if (_character.Skills_with_points.IndexOf(skill) + 1 < 19)
                            {
                                skill_coloumn_num = 19;
                                skill_score_coloumn_num = 21;
                            }
                            else if (_character.Skills_with_points.IndexOf(skill) + 1 == 19)
                            {
                                row_index = 5;
                                skill_coloumn_num = 22;
                                skill_score_coloumn_num = 23;
                            }
                            else
                            {
                                skill_coloumn_num = 22;
                                skill_score_coloumn_num = 23;
                            }
                            Character_card.Cells[row_index, skill_coloumn_num].Value = skill.Name;
                            Character_card.Cells[row_index, skill_score_coloumn_num].Value = skill.Score;
                            row_index = (byte)(row_index + 1);
                        }
                    }

                    // Заполняем поля навыков Силы
                    if (_character.Force_skills_with_points.Count > 0)
                    {
                        skill_coloumn_num = 17;
                        skill_score_coloumn_num = 18;
                        row_index = 5;
                        foreach (Force_skill_class skill in _character.Force_skills_with_points)
                        {
                            Character_card.Cells[row_index, skill_coloumn_num].Value = skill.Name;
                            Character_card.Cells[row_index, skill_score_coloumn_num].Value = skill.Score;
                            row_index = (byte)(row_index + 1);
                        }
                    }

                    // Заполняем положительные особенности
                    if (_character.Positive_features_with_points.Count > 0)
                    {
                        skill_coloumn_num = 17;
                        row_index = 26;
                        foreach (All_feature_template feature in _character.Positive_features_with_points)
                        {
                            Character_card.Cells[row_index, skill_coloumn_num].Value = feature.Name;
                            row_index = (byte)(row_index + 1);
                        }
                    }

                    // Заполняем отрицательные особенности
                    if (_character.Negative_features_with_points.Count > 0)
                    {
                        skill_coloumn_num = 20;
                        row_index = 26;
                        foreach (All_feature_template feature in _character.Negative_features_with_points)
                        {
                            Character_card.Cells[row_index, skill_coloumn_num].Value = feature.Name;
                            row_index = (byte)(row_index + 1);
                        }
                    }

                    package.Save();
                    //package.SaveAs(character_directory + $"\\{_character.Name}.character");

                    // Концертируем карточку из формата экселя в формат PDF
                    Workbook workbook = new Workbook();
                    workbook.LoadFromFile(character_file);
                    workbook.SaveToFile(character_directory + $"\\{_character.Name}.pdf", Spire.Xls.FileFormat.PDF);
                }
                File.Copy(character_directory + $"\\{_character.Name}.xlsx", character_directory + $"\\{_character.Name}.character",true);
            }
            else
            {
                //error_msg = "Создание анкеты персонажа невозможно! Отсутствует шаблон анкеты.";
                throw new Exception("Отсутствует шаблон карточки персонажа! Сохранение невозможно!");
            }
        }
    }
}
