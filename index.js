const cheerio = require('cheerio');
const axios = require('axios');
const ExcelJS = require('exceljs');

const workbook = new ExcelJS.Workbook();
const worksheet = workbook.addWorksheet('Player Stats');
const header = [
  'Team ID', 'Team Name', 'Player Name', 'Games Played', 
  'Minutes Played', 'Points per Game', 'Rebounds per Game',
  'Assists per Game', 'Steals per Game', 'Blocks per Game',
  'Field Goal Percentage', 'Three-Point Percentage', 'Free Throw Percentage'
];
worksheet.addRow(header);

const pages = [
  2, 150, 248, 57, 2633, 333, 127, 2599, 2641, 66, 96, 275, 245, 2509, 120, 
  12, 130, 228, 2483, 235, 252, 356, 142, 145, 26, 269, 2608, 2305, 97, 2250, 
  41, 344, 156, 61, 239, 201, 8, 167, 238, 328, 251, 2752, 21, 2181, 2670, 153, 
  28, 36, 2377, 2335, 43, 2272, 2006, 2253, 288, 2653, 350, 149, 2523, 2747, 
  2437, 2803, 2450, 2565, 44, 116, 2011, 2598
];

async function scrape() {
  for (const page of pages) {
    try {
      const response = await axios.get(`https://www.espn.com/mens-college-basketball/team/stats/_/id/${page}`);
      const $ = cheerio.load(response.data);

      const team_name = $('h1.headline.headline__h1.dib').text().replace(/[\d]|[\-]|(\Stats)/g, '');
      const left_table = $('table:eq(0)');
      const right_table = $('table:eq(1)');
      const left_rows = left_table.find('tr').slice(1, -1);
      const right_rows = right_table.find('tr').slice(1, -1);

      const stats = [];
      left_rows.each((index, row) => {
        const left_cols = $(row).find('td');
        const right_cols = $(right_rows[index]).find('td');
        const player_name = left_cols.eq(0).find('a').text();
        const games_played = parseFloat(right_cols.eq(0).text());
        const minutes_played = parseFloat(right_cols.eq(1).text());
        const points_per_game = parseFloat(right_cols.eq(2).text());
        const rebounds_per_game = parseFloat(right_cols.eq(3).text());
        const assists_per_game = parseFloat(right_cols.eq(4).text());
        const steals_per_game = parseFloat(right_cols.eq(5).text());
        const blocks_per_game = parseFloat(right_cols.eq(6).text());
        const field_goal_percentage = parseFloat(right_cols.eq(8).text());
        const free_throw_percentage = parseFloat(right_cols.eq(9).text());
        const three_point_percentage = parseFloat(right_cols.eq(10).text());

        stats.push([
          page, team_name, player_name, games_played, minutes_played, points_per_game, 
          rebounds_per_game, assists_per_game, steals_per_game, blocks_per_game, 
          field_goal_percentage + '%', three_point_percentage + '%', free_throw_percentage + '%'
        ]);
      });

      // Append data to worksheet in order
      stats.forEach(row => worksheet.addRow(row));

    } catch (error) {
      console.error(`Error scraping page ${page}:`, error.message);
    }
  }

  // Write Excel file after all pages are processed
  await workbook.xlsx.writeFile('player_stats.xlsx');
  console.log('Excel file saved as player_stats.xlsx');
}

scrape();
