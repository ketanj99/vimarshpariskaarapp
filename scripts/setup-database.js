const { Pool } = require('pg');
const fs = require('fs');
const path = require('path');
require('dotenv').config({ path: path.join(__dirname, '..', '.env.local') });

async function setupDatabase() {
  const pool = new Pool({
    host: process.env.DB_HOST || 'localhost',
    port: parseInt(process.env.DB_PORT || '5432'),
    database: process.env.DB_NAME || 'vimarshbooks',
    user: process.env.DB_USER || 'postgres',
    password: process.env.DB_PASSWORD || '',
  });

  try {
    console.log('🔌 Connecting to PostgreSQL...');
    const client = await pool.connect();
    console.log('✅ Connected to PostgreSQL');

    // Run init schema
    console.log('\n📋 Running schema initialization...');
    const initSql = fs.readFileSync(
      path.join(__dirname, '..', 'migrations', '000_init_schema.sql'),
      'utf8'
    );
    await client.query(initSql);
    console.log('✅ Schema initialized');

    // Run migrations
    console.log('\n📋 Running migrations...');
    const migrationSql = fs.readFileSync(
      path.join(__dirname, '..', 'migrations', '001_add_app_to_transactions.sql'),
      'utf8'
    );
    await client.query(migrationSql);
    console.log('✅ Migrations completed');

    // Verify tables
    console.log('\n🔍 Verifying tables...');
    const result = await client.query(`
      SELECT table_name 
      FROM information_schema.tables 
      WHERE table_schema = 'vimars' 
      ORDER BY table_name
    `);

    console.log('✅ Tables found:');
    result.rows.forEach((row) => {
      console.log('   - ' + row.table_name);
    });

    client.release();
    await pool.end();

    console.log('\n✨ Database setup completed successfully!');
    console.log('\n🚀 You can now start the application with: npm run dev');
  } catch (error) {
    console.error('\n❌ Error setting up database:', error.message);
    console.error('\nPlease check:');
    console.error('  1. PostgreSQL is running');
    console.error('  2. Database credentials in .env.local are correct');
    console.error('  3. Database "vimarshbooks" exists');
    process.exit(1);
  }
}

setupDatabase();
