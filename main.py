import xml.etree.ElementTree as ET
import numpy as np
import xlsxwriter

class Team:

    def __init__(self, xml, squad_details):

        self.simple_stat_lines = {                    
            'apps': 'appearances',
            'starts': 'starts',
            'sub_on': 'sub appearances',
            'minutes': 'minutes played',
            'assists': 'assists',
            'involvements': 'goal involvements',
            'chances_created': 'chances created',
            'aerials_won': 'aerial duels won',
            'shots_on_target': 'shots on target',
            'blocks': 'blocks',
            'recoveries': 'possession won',
            'passes_opp_half': 'successful passes in opp. half',
            'dribbles': 'successful dribbles',
            'interceptions': 'interceptions',
            'successful_passes': 'successful passes',
            'long_passes': 'successful long passes',
            'ground_duels': 'ground duels won',
            'clearances': 'clearances',
            'tackles': 'tackles',
            'through_balls': 'through balls',
            'winners': 'winning goals'}
        
        self.complex_stat_lines = {
            'sub_off': 'Subbed off most times for {} this season ({} times)',
            'goals': '{}\'s top scorer this season with {} goals'}
        
        self.comp_names = {'English Premier League': 'Premier League'}

        self.squad_details_tree = ET.parse(squad_details)
        self.squad_details_root = self.squad_details_tree.getroot()
        self.squad_details = self.get_soccerdocument()
        
        self.tree = ET.parse(xml)
        self.root = self.tree.getroot()

        self.season = self.format_season()
        self.team_id = self.root[0].attrib['id']
        self.team_name = self.root[0].attrib['name']
        
        if self.root.attrib['competition_name'] in self.comp_names:
            self.comp = self.comp_names[self.root.attrib['competition_name']]
        else:
            self.comp = self.root.attrib['competition_name']

        self.header_text = '{} in {} {}'.format(self.team_name,
                                                self.season,
                                                self.comp)


        self.squad_root = self.root.findall('.//Player')

        self.squad = self.build_squad()

        for stat in self.simple_stat_lines:
            max_value, max_key = self.find_stat_leader(stat)
            stat_line = self.build_stat_line(stat, max_value)

            for player in self.squad:
                if max_key == player.opta_id:
                    player.stat_lines.append(stat_line)

    def get_soccerdocument(self):
    
        for element in self.squad_details_root:
            if element.tag == 'SoccerDocument':
                return element

    def minute_threshold(self, percentile):
        self.squad_minutes = [player.minutes for player in self.squad if player.minutes > 0]
        self.treshold = int(np.percentile(self.squad_minutes, percentile))

    def format_season(self):
        
        season_sting = self.root.attrib['season_name'][7:]
        season_sting = season_sting.replace('/', '-')

        return season_sting
    
    def build_squad(self, sortby='position'):

        squadlist = []

        for player in self.squad_root:

            player = self.Player(player, squad_details=self.squad_details)
            squadlist.append(player)

        if sortby == 'position':

            squadlist.sort(key=lambda player: player.number)
            squadlist.sort(key=lambda player: player.basic_position_key)

        return squadlist
    
    def find_stat_leader(self, stat):

        stat_values = {}

        for player in self.squad:

            stat_values[player.opta_id] = getattr(player, stat)

        max_value = max(stat_values.values())
        max_key = [key for key, value in stat_values.items() if value == max_value][0]

        return max_value, max_key
    
    def build_stat_line(self, stat, value):

        if stat in self.complex_stat_lines:
            stat_line = self.complex_stat_lines[stat].format(self.team_name, value)
        elif stat in self.simple_stat_lines:
            stat_line = 'Most {} ({}) for {} this season'.format(self.simple_stat_lines[stat], value, self.team_name)
    
        return stat_line

    class Player:

        def __init__(self, player, squad_details):

            self.position_keys = {'Goalkeeper': 1,
                                  'Defender': 2,
                                  'Midfielder': 3,
                                  'Forward': 4}

            self.first_name = player.attrib['first_name']
            self.surname = player.attrib['last_name']

            try:
                self.known_name = player.attrib['known_name']
            except KeyError:
                self.known_name = None

            self.number = int(player.attrib['shirtNumber'])
            self.basic_position = player.attrib['position']

            self.basic_position_key = self.position_keys[self.basic_position]

            self.opta_id = player.attrib['player_id']

            self.player_stats = player.findall('.//Stat')

            self.apps = self.find_value('Appearances')
            self.starts = self.find_value('Starts')
            self.sub_on = self.find_value('Substitute On')
            self.sub_off = self.find_value('Substitute Off')
            self.minutes = self.find_value('Time Played')

            self.nation = self.get_details(self.opta_id, 'first_nationality', squad_details)

            if self.nation == None:
                try:
                    self.nation = self.get_details(self.opta_id, 'country', squad_details)
                except ValueError:
                    self.nation = ''

            self.preferred_foot = self.get_details(self.opta_id, 'preferred_foot', squad_details)
            self.shirt_number = self.get_details(self.opta_id, 'jersey_num', squad_details)

            if self.shirt_number == None:
                self.shirt_number = 'xx'

            self.height_cm = self.get_details(self.opta_id, 'height', squad_details)

            try:
                self.height_cm = int(self.height_cm)
                self.height_string = self.cm_to_feet()
            except (ValueError, TypeError):
                self.height_string = 'Unknown'

            self.detailled_position = self.get_details(self.opta_id, 'real_position', squad_details)

            self.goals = (self.find_value('Goals'))
            self.assists = (self.find_value('Goal Assists'))
            self.involvements = self.goals + self.assists

            self.chances_created = self.find_value('Key Passes (Attempt Assists)')
            self.aerials_won = self.find_value('Aerial Duels won')
            self.shots_on_target = self.find_value('Shots On Target ( inc goals )')
            self.blocks = self.find_value('Blocks')
            self.recoveries = self.find_value('Recoveries')
            self.passes_opp_half = self.find_value('Successful Passes Opposition Half')
            self.dribbles = self.find_value('Successful Dribbles')
            self.interceptions = self.find_value('Interceptions')
            self.successful_passes = self.find_value('Total Successful Passes ( Excl Crosses & Corners ) ')
            self.long_passes = self.find_value('Successful Long Passes')
            self.ground_duels = self.find_value('Ground Duels won')
            self.clearances = self.find_value('Total Clearances')
            self.tackles = self.find_value('Total Tackles')
            self.through_balls = self.find_value('Through balls')
            self.winners = self.find_value('Winning Goal')

            self.stat_lines = []

            self.pen_pic = self.output_penpic()

            if self.known_name:
                self.header = '{}. {}      {}'.format(self.number, self.known_name, self.nation)
            else:
                self.header = '{}. {} {}      {}'.format(self.number, self.first_name, self.surname, self.nation)

            self.app_line = ('Appearances: {}       Starts: {}       Sub on:  {}       Sub off: {}       Minutes: {}'.format(self.apps, 
                                                                                                            self.starts, 
                                                                                                            self.sub_on, 
                                                                                                            self.sub_off, 
                                                                                                            self.minutes))
            
            if self.basic_position == 'Goalkeeper':

                self.clean_sheets = self.find_value('Clean Sheets')
                self.pens_faced = self.find_value('Penalties Faced')
                self.pens_saved = self.find_value('Penalties Saved')

                self.ga_line = ('Clean sheets: {}       Penalties faced: {}       Penalties saved: {}'.format(self.clean_sheets,
                                                                                                      self.pens_faced,
                                                                                                      self.pens_saved))
                
            else:

                self.ga_line = ('Goals: {}       Assists: {}       Goal Involvements: {}'.format(self.goals,
                                                                                         self.assists,
                                                                                         self.involvements))
            
        def cm_to_feet(self):
    
            inches = self.height_cm / 2.54
            feet_and_inches = inches / 12
    
            inch_remainder = feet_and_inches - int(feet_and_inches)
            
            inch_remainder = int(inch_remainder * 12)
            
            if inch_remainder >= 1:
                return '{} ft {} in'.format(int(feet_and_inches), inch_remainder)
            else:
                return '{} ft'.format(int(feet_and_inches))
        
        def find_value(self, value):

            for stat in self.player_stats:

                if stat.attrib['name'] == value:
                    return int(stat.text)
            return 0
        
        def get_details(self, player_id, attribute, squad_details):
            
            if type(player_id) != str:
                player_id = str(player_id)

            player_id = 'p' + player_id

            for team in squad_details:
                for player in team:
                    if 'uID' in player.attrib:
                        if player.attrib['uID'] == player_id:
                            for attributes in player:
                                if 'Type' in attributes.attrib and attributes.attrib['Type'] == attribute:
                                    return attributes.text
                                
            return None
        
        def output_penpic(self):

            name = self.known_name if self.known_name else '{} {}'.format(self.first_name, self.surname)
            topline = '{}. {}        {}'.format(self.shirt_number, name, self.nation) if self.nation else '{}. {}'.format(self.number, name) 

            line_two = 'Position: {}    Preferred foot: {}    Height: {}'.format(self.detailled_position, self.preferred_foot, self.height_string)

            app_line = 'Apps: {}   Starts: {}   Sub on: {}   Sub off: {}   Minutes: {}'.format(self.apps,
                                                                                               self.starts,
                                                                                               self.sub_on,
                                                                                               self.sub_off,
                                                                                               self.minutes)
            
            involvments_line = 'Goals: {}      Assists: {}      Goal involvements: {}'.format(self.goals,
                                                                                              self.assists,
                                                                                              self.involvements)

            print(topline)
            print('')
            print(line_two)
            print(app_line)
            print(involvments_line)

            print('')

            for line in self.stat_lines:
                print(line)

            print('')
            print('')

            pen_pic = [topline,
                       line_two,
                       app_line,
                       involvments_line,
                       self.stat_lines]
            
            return pen_pic
        
def write_excel(filename, stats):

    team_header = '{} squad'.format(stats.team_name)

    workbook = xlsxwriter.Workbook(filename)
    worksheet = workbook.add_worksheet(name='Leicester pen pics')
    worksheet.fit_to_pages(1,99)

    sheet_width = 9

    header_format = workbook.add_format({'bold': True,
                                         'border': True,
                                         'font': 'calibri',
                                         'size': 22,
                                         'align': 'center',
                                         'fg_color': 'blue',
                                         'valign': 'vcenter',
                                         'font_color': 'white'})
    
    player_name_format = workbook.add_format({'bold': True,
                                              'border': True,
                                              'font': 'calibri',
                                              'size': 14,
                                              'fg_color': 'yellow',
                                              'valign': 'top',
                                              'text_wrap': True})
    
    stat_format = workbook.add_format({'bold': 1,
                                       'size': 12})
    



    worksheet.merge_range(first_row=0, last_row=4, first_col=0, last_col=sheet_width, data=team_header, cell_format=header_format)
    
    worksheet.insert_image('A2', 'Leicester.png', {'object_position': 1})
    worksheet.insert_image('J2', 'Leicester.png', {'object_position': 1})

    start_row = 6

    tracking_row = start_row

    for player in stats.squad:
        worksheet.merge_range(first_row=tracking_row,
                              last_row=tracking_row,
                              first_col=0,
                              last_col=sheet_width,
                              data=player.header,
                              cell_format=player_name_format)
        
        print(player.header)
        
        tracking_row += 1

        if player.detailled_position == None:
            player.detailled_position = player.basic_position
        
        worksheet.write(tracking_row, 0, player.detailled_position, stat_format)

        if player.preferred_foot != None:
            worksheet.write(tracking_row, 4, 'Foot: {}'.format(player.preferred_foot), stat_format)

        if player.height_string != 'Unknown':
            worksheet.write(tracking_row, 7, 'Height: ' + player.height_string, stat_format)

        tracking_row += 2

        worksheet.write(tracking_row, 0, '{} {}:'.format(stats.season, stats.comp), stat_format)

        tracking_row += 2

        worksheet.write(tracking_row, 0, 'Appearances:',stat_format)
        worksheet.write(tracking_row, 2, player.apps)

        if player.apps > 0:

            worksheet.write(tracking_row+1, 0, 'Starts:',stat_format)
            worksheet.write(tracking_row+2, 0, 'Subbed on:',stat_format)
            worksheet.write(tracking_row+3, 0, 'Subbed off:',stat_format)
            worksheet.write(tracking_row+4, 0, 'Minutes played:',stat_format)


            worksheet.write(tracking_row+1, 2, player.starts)
            worksheet.write(tracking_row+2, 2, player.sub_on)
            worksheet.write(tracking_row+3, 2, player.sub_off)
            worksheet.write(tracking_row+4, 2, player.minutes)

            if player.basic_position == 'Goalkeeper':
                worksheet.write(tracking_row, 4, 'Clean sheets:', stat_format)
                worksheet.write(tracking_row+2, 4, 'Penalties faced:', stat_format)
                worksheet.write(tracking_row+3, 4, 'Penalties saved:', stat_format)

                worksheet.write(tracking_row, 6, player.clean_sheets)

                if player.pens_faced > 0:

                    worksheet.write(tracking_row+2, 6, player.pens_faced)
                    worksheet.write(tracking_row+3, 6, player.pens_saved)

            else:
                worksheet.write(tracking_row, 4, 'Goals:', stat_format)
                worksheet.write(tracking_row+1, 4, 'Assists:', stat_format)
                worksheet.write(tracking_row+3, 4, 'Involvements:', stat_format)

                worksheet.write(tracking_row, 6, player.goals)
                worksheet.write(tracking_row+1, 6, player.assists)
                worksheet.write(tracking_row+3, 6, player.involvements)

            tracking_row += 6


#        if player.apps > 0:
#            if player.basic_position == 'Goalkeeper':
#                worksheet.write(tracking_row, 0, 'Clean sheets: {}'.format(player.clean_sheets))
#            else:
#                worksheet.write(tracking_row, 0, 'Goals: {}'.format(player.goals))
#                worksheet.write(tracking_row, 2, 'Assists: {}'.format(player.assists))
#                worksheet.write(tracking_row, 4, 'Goal involvements: {}'.format(player.involvements))
        
        else:
            tracking_row += 2

        if len(player.stat_lines) > 0:
            worksheet.write(tracking_row, 0, 'Season stats:', stat_format)
            tracking_row += 1

        for line in player.stat_lines:
            worksheet.write(tracking_row, 0, line)
            tracking_row += 1
        
        if len(player.stat_lines) > 0:
            tracking_row += 1







    workbook.close()


if __name__ == '__main__':

    data = './/sample.xml'
    stats = Team(data, squad_details='.//2021_squads.xml')

    a = stats.minute_threshold(33)
    
    playergoals = sorted([player.goals for player in stats.squad], reverse=True)
    playerassists = sorted([player.assists for player in stats.squad], reverse=True) 

    write_excel('test.xlsx', stats)


#    for player in stats.squad:

#        player.output_penpic()
        
#        if player.goals == max(playergoals):
#            print('{} {} is top scorer with {} goals'.format(player.first_name, player.surname, player.goals))
#        if player.goals == playergoals[1]:
#            print('{} {} is 2nd-top scorer with {} goals'.format(player.first_name, player.surname, player.goals))

#        if player.assists == max(playerassists):
#            print('{} {} has most assists with {}'.format(player.first_name, player.surname, player.assists))
#            leadername = player.first_name + ' ' + player.surname
#        if player.assists == playerassists[2]:
#            print('{} {} has 2nd-most assists ({}) behind {}'.format(player.first_name, player.surname, player.assists, leadername))