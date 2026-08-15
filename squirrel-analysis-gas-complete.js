// ========================================
// Squirrel Designer - Google Apps Script Backend
// 佣金分开储存版本 - 兼容前台系统
// Quote Analysis 报价分析系统 整合版
// ========================================

const CONFIG = {
  MAIN_FOLDER: 'Squirrel Designer',
  USERS_FILE: 'users.json',
  PRODUCTS_FILE: 'products.json',
  ACTIVITY_FILE: 'activity_logs.json',
  ADMIN_LOG_FILE: 'admin_logs.json',
  COMMISSION_FOLDER: 'Commission'
};

let activityCache = null;
let adminLogCache = null;

let mainFolder = null;
let adminFolder = null;
let analysisFolder = null;
let commissionFolder = null;
let usersCache = null;
let productsCache = null;

// 2026-08-15 性能优化：缓存层（用 CacheService 持久化跨调用）
// 教训：模块级 let 变量在 GAS Web App 跨调用会被重置（新脚本上下文），不能用！
let folderListCache = null;     // 留作 fallback（单次调用内仍可省重复 set）
let userFolderCache = {};
let checkQuoteListCache = null;
const CACHE_TTL_FOLDERS = 30;   // CacheService TTL 单位是秒
const CACHE_TTL_LIST = 10;
const CACHE_TTL_USER_FOLDER = 30;

// 跨调用持久化缓存（CacheService.getScriptCache — 6h max, 100KB max）
function cacheGet(key) {
  try { const v = CacheService.getScriptCache().get(key); return v ? JSON.parse(v) : null; }
  catch (e) { return null; }
}
function cachePut(key, value, ttlSec) {
  try { CacheService.getScriptCache().put(key, JSON.stringify(value), ttlSec); } catch (e) {}
}
function cacheRemove(key) {
  try { CacheService.getScriptCache().remove(key); } catch (e) {}
}

// ==================== 初始化 ====================

function initializeFolders() {
  try {
    const rootFolders = DriveApp.getRootFolder().getFoldersByName(CONFIG.MAIN_FOLDER);
    mainFolder = rootFolders.hasNext() ? rootFolders.next() : DriveApp.getRootFolder().createFolder(CONFIG.MAIN_FOLDER);
    
    const adminFolders = mainFolder.getFoldersByName('admin');
    adminFolder = adminFolders.hasNext() ? adminFolders.next() : mainFolder.createFolder('admin');
    
    const analysisFolders = mainFolder.getFoldersByName('squirrel analysis');
    analysisFolder = analysisFolders.hasNext() ? analysisFolders.next() : mainFolder.createFolder('squirrel analysis');
    
    const commissionFolders = mainFolder.getFoldersByName(CONFIG.COMMISSION_FOLDER);
    commissionFolder = commissionFolders.hasNext() ? commissionFolders.next() : mainFolder.createFolder(CONFIG.COMMISSION_FOLDER);
    
    const userFiles = adminFolder.getFilesByName(CONFIG.USERS_FILE);
    if (!userFiles.hasNext()) {
      usersCache = { users: [{ 
        username: 'admin', 
        displayName: 'Administrator', 
        password: 'admin1234', 
        isAdmin: true, 
        isActive: true, 
        departments: ['admin'],
        createdAt: new Date().toISOString(),
        lastLogin: null
      }] };
      adminFolder.createFile(CONFIG.USERS_FILE, JSON.stringify(usersCache));
    } else {
      const file = userFiles.next();
      usersCache = JSON.parse(file.getBlob().getDataAsString());
    }
    
    // Initialize activity logs
    const activityFiles = adminFolder.getFilesByName(CONFIG.ACTIVITY_FILE);
    if (!activityFiles.hasNext()) {
      activityCache = { activities: [] };
      adminFolder.createFile(CONFIG.ACTIVITY_FILE, JSON.stringify(activityCache));
    } else {
      const file = activityFiles.next();
      activityCache = JSON.parse(file.getBlob().getDataAsString());
    }
    
    // Initialize admin logs
    const adminLogFiles = adminFolder.getFilesByName(CONFIG.ADMIN_LOG_FILE);
    if (!adminLogFiles.hasNext()) {
      adminLogCache = { logs: [] };
      adminFolder.createFile(CONFIG.ADMIN_LOG_FILE, JSON.stringify(adminLogCache));
    } else {
      const file = adminLogFiles.next();
      adminLogCache = JSON.parse(file.getBlob().getDataAsString());
    }
    
    return { success: true, message: 'Initialized successfully' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// ==================== 用户管理 ====================

function login(username, password) {
  try {
    initializeFolders();
    const user = usersCache.users.find(u => u.username === username && u.password === password);
    
    if (!user) {
      logActivity(username, 'login_failed', null, 'Invalid username or password');
      return { success: false, message: 'Invalid username or password' };
    }
    if (!user.isActive) {
      logActivity(username, 'login_failed', null, 'Account is frozen');
      return { success: false, message: 'Account is frozen' };
    }
    
    user.lastLogin = new Date().toISOString();
    saveUsers();
    logActivity(username, 'login', null, 'Logged in successfully');
    
    let userFolder;
    const folders = mainFolder.getFoldersByName(username);
    userFolder = folders.hasNext() ? folders.next() : mainFolder.createFolder(username);
    
    return { 
      success: true, 
      user: { 
        username: user.username, 
        displayName: user.displayName || user.username,
        isAdmin: user.isAdmin || false, 
        isAgent: user.isAgent || false,
        departments: user.departments || []
      } 
    };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function saveUsers() {
  const files = adminFolder.getFilesByName(CONFIG.USERS_FILE);
  while (files.hasNext()) files.next().setTrashed(true);
  adminFolder.createFile(CONFIG.USERS_FILE, JSON.stringify(usersCache, null, 2));
}

function getAllUsers() {
  try {
    initializeFolders();
    return { success: true, users: usersCache.users.map(u => ({ username: u.username, isAdmin: u.isAdmin || false, isActive: u.isActive, isAgent: u.isAgent || false, createdAt: u.createdAt })) };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function addUser(username, password, isAgent) {
  try {
    initializeFolders();
    if (usersCache.users.find(u => u.username === username)) return { success: false, message: 'Username exists' };
    
    usersCache.users.push({ username, password: password || 'password123', isAdmin: false, isActive: true, isAgent: isAgent || false, createdAt: new Date().toISOString() });
    saveUsers();
    mainFolder.createFolder(username);
    
    return { success: true, message: 'User added' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function toggleUserStatus(username) {
  try {
    initializeFolders();
    const user = usersCache.users.find(u => u.username === username);
    if (!user) return { success: false, message: 'User not found' };
    
    user.isActive = !user.isActive;
    saveUsers();
    return { success: true, user };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function deleteUser(username) {
  try {
    initializeFolders();
    if (username === 'admin') return { success: false, message: 'Cannot delete admin' };
    
    const index = usersCache.users.findIndex(u => u.username === username);
    if (index === -1) return { success: false, message: 'User not found' };
    
    usersCache.users.splice(index, 1);
    saveUsers();
    return { success: true };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// ==================== 活动日志 ====================

function logActivity(username, action, projNo, details) {
  try {
    if (!activityCache) initializeFolders();
    const activity = {
      id: Date.now(),
      username: username,
      action: action,
      projNo: projNo || null,
      details: details || '',
      timestamp: new Date().toISOString()
    };
    activityCache.activities.push(activity);
    saveActivity();
  } catch (e) {
    console.error('Failed to log activity:', e);
  }
}

function saveActivity() {
  const files = adminFolder.getFilesByName(CONFIG.ACTIVITY_FILE);
  while (files.hasNext()) files.next().setTrashed(true);
  adminFolder.createFile(CONFIG.ACTIVITY_FILE, JSON.stringify(activityCache, null, 2));
}

function getUserActivities(username) {
  try {
    initializeFolders();
    const userActivities = activityCache.activities
      .filter(a => a.username === username)
      .sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
    return { success: true, activities: userActivities };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function getAllActivities() {
  try {
    initializeFolders();
    const activities = activityCache.activities
      .sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
    return { success: true, activities: activities };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// ==================== 管理员功能 ====================

function adminLogin(username, password) {
  try {
    initializeFolders();
    if (username !== 'admin') return { success: false, message: 'Not an admin account' };
    
    const admin = usersCache.users.find(u => u.username === 'admin' && u.password === password);
    if (!admin) return { success: false, message: 'Invalid password' };
    
    admin.lastLogin = new Date().toISOString();
    saveUsers();
    logActivity('admin', 'admin_login', null, 'Admin logged in');
    
    return { success: true, message: 'Admin login successful' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function changeAdminPassword(oldPassword, newPassword) {
  try {
    initializeFolders();
    const admin = usersCache.users.find(u => u.username === 'admin');
    if (!admin) return { success: false, message: 'Admin not found' };
    if (admin.password !== oldPassword) return { success: false, message: 'Current password is incorrect' };
    
    admin.password = newPassword;
    saveUsers();
    logActivity('admin', 'change_password', null, 'Admin changed password');
    
    return { success: true, message: 'Password changed successfully' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function getUserByUsername(username) {
  try {
    initializeFolders();
    const user = usersCache.users.find(u => u.username === username);
    if (!user) return { success: false, message: 'User not found' };
    
    return { 
      success: true, 
      user: {
        username: user.username,
        displayName: user.displayName || user.username,
        password: user.password,
        isAdmin: user.isAdmin || false,
        isActive: user.isActive,
        departments: user.departments || [],
        createdAt: user.createdAt,
        lastLogin: user.lastLogin
      }
    };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function updateUser(username, updates) {
  try {
    initializeFolders();
    const user = usersCache.users.find(u => u.username === username);
    if (!user) return { success: false, message: 'User not found' };
    
    if (updates.displayName !== undefined) user.displayName = updates.displayName;
    if (updates.password !== undefined) user.password = updates.password;
    if (updates.departments !== undefined) user.departments = updates.departments;
    if (updates.isActive !== undefined) user.isActive = updates.isActive;
    
    saveUsers();
    logActivity('admin', 'update_user', null, 'Updated user: ' + username);
    
    return { success: true, message: 'User updated successfully' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function getUserStatistics(username) {
  try {
    initializeFolders();
    const userFolder = getUserFolder(username);
    const files = userFolder.getFiles();
    
    let totalQuotes = 0;
    let totalSales = 0;
    let totalDebt = 0;
    let quotes = [];
    
    while (files.hasNext()) {
      const file = files.next();
      const name = file.getName();
      if (name.endsWith('.json') && !name.startsWith('final_')) {
        try {
          const quote = JSON.parse(file.getBlob().getDataAsString());
          totalQuotes++;
          const grandTotal = quote.total || quote.grandTotal || 0;
          let paidAmount = 0;
          if (quote.depositRecords && quote.depositRecords.length > 0) {
            paidAmount = quote.depositRecords.reduce((sum, dep) => sum + (dep.amount || 0), 0);
          }
          const debt = grandTotal - paidAmount;
          
          totalSales += grandTotal;
          totalDebt += debt;
          quotes.push({
            projNo: quote.projNo || name.replace('.json', ''),
            customerName: quote.customerName || '',
            grandTotal: grandTotal,
            paidAmount: paidAmount,
            debt: debt,
            status: quote.status || 'active',
            designerName: quote.designerName || '',
            subSalesperson: quote.subSalesperson || '',
            measurementEngineer: quote.measurementEngineer || ''
          });
        } catch (e) {}
      }
    }
    
    return {
      success: true,
      statistics: {
        totalQuotes: totalQuotes,
        totalSales: totalSales,
        totalDebt: totalDebt
      },
      quotes: quotes
    };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function getAdminStats() {
  try {
    initializeFolders();
    
    let totalUsers = usersCache.users.length;
    let totalSales = 0;
    let totalDebt = 0;
    
    const subFolders = mainFolder.getFolders();
    while (subFolders.hasNext()) {
      const folder = subFolders.next();
      if (folder.getName() === 'admin' || folder.getName() === 'squirrel analysis') continue;
      
      const files = folder.getFiles();
      while (files.hasNext()) {
        const file = files.next();
        const name = file.getName();
        if (name.endsWith('.json') && !name.startsWith('final_')) {
          try {
            const quote = JSON.parse(file.getBlob().getDataAsString());
            const grandTotal = quote.total || quote.grandTotal || 0;
            let paidAmount = 0;
            if (quote.depositRecords && quote.depositRecords.length > 0) {
              paidAmount = quote.depositRecords.reduce((sum, dep) => sum + (dep.amount || 0), 0);
            }
            const debt = grandTotal - paidAmount;
            
            totalSales += grandTotal;
            totalDebt += debt;
          } catch (e) {}
        }
      }
    }
    
    return {
      success: true,
      stats: {
        totalUsers: totalUsers,
        totalSales: totalSales,
        totalDebt: totalDebt
      }
    };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function getDetailedUsers() {
  try {
    initializeFolders();
    const users = usersCache.users.map(u => ({
      username: u.username,
      displayName: u.displayName || u.username,
      isAdmin: u.isAdmin || false,
      isActive: u.isActive,
      departments: u.departments || [],
      createdAt: u.createdAt,
      lastLogin: u.lastLogin
    }));
    
    return { success: true, users: users };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function adminAddUser(username, password, displayName, departments) {
  try {
    initializeFolders();
    if (usersCache.users.find(u => u.username === username)) {
      return { success: false, message: 'Username already exists' };
    }
    
    const newUser = {
      username: username,
      displayName: displayName || username,
      password: password || 'password123',
      isAdmin: false,
      isActive: true,
      departments: departments || [],
      createdAt: new Date().toISOString(),
      lastLogin: null
    };
    
    usersCache.users.push(newUser);
    saveUsers();
    mainFolder.createFolder(username);
    
    logActivity('admin', 'add_user', null, 'Added new user: ' + username);

    return { success: true, message: 'User added successfully' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function adminDeleteUser(username) {
  try {
    initializeFolders();
    if (username === 'admin') {
      return { success: false, message: 'Cannot delete admin account' };
    }
    
    const index = usersCache.users.findIndex(u => u.username === username);
    if (index === -1) {
      return { success: false, message: 'User not found' };
    }
    
    usersCache.users.splice(index, 1);
    saveUsers();
    
    logActivity('admin', 'delete_user', null, 'Deleted user: ' + username);
    
    return { success: true, message: 'User deleted successfully' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function adminToggleUserStatus(username) {
  try {
    initializeFolders();
    const user = usersCache.users.find(u => u.username === username);
    if (!user) {
      return { success: false, message: 'User not found' };
    }
    
    if (username === 'admin') {
      return { success: false, message: 'Cannot freeze admin account' };
    }
    
    user.isActive = !user.isActive;
    saveUsers();
    
    const action = user.isActive ? 'unfreeze_user' : 'freeze_user';
    const details = user.isActive ? 'Unfroze user: ' + username : 'Froze user: ' + username;
    logActivity('admin', action, null, details);
    
    return { success: true, message: user.isActive ? 'User activated' : 'User frozen', isActive: user.isActive };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// ==================== 报价单管理 ====================

function getUserFolder(username) {
  const folders = mainFolder.getFoldersByName(username);
  return folders.hasNext() ? folders.next() : mainFolder.createFolder(username);
}

function saveQuote(username, projNo, quoteData) {
  try {
    initializeFolders();
    const userFolder = getUserFolder(username);
    const fileName = projNo + '.json';
    
    const files = userFolder.getFilesByName(fileName);
    while (files.hasNext()) files.next().setTrashed(true);
    
    userFolder.createFile(fileName, JSON.stringify(quoteData, null, 2), 'application/json');
    logActivity(username, 'save_quote', projNo, 'Saved quote: ' + (quoteData.customerName || projNo));
    return { success: true, message: 'Quote saved' };
} catch (error) {
    return { success: false, message: error.toString() };
  }
}

function getQuoteList(username) {
  try {
    initializeFolders();
    const userFolder = getUserFolder(username);
    const files = userFolder.getFiles();
    const quotes = [];
    
    while (files.hasNext()) {
      const file = files.next();
      if (file.getName().endsWith('.json') && !file.getName().startsWith('final_')) {
        try {
          const data = JSON.parse(file.getBlob().getDataAsString());
          quotes.push({
            id: data.id || file.getName().replace('.json', ''),
            projNo: data.projNo || file.getName().replace('.json', ''),
            customerName: data.customerName || '',
            customerIC: data.customerIC || '',
            customerContact: data.customerContact || '',
            customerAddress: data.customerAddress || '',
            salesperson: data.salesperson || '',
            salespersonContact: data.salespersonContact || '',
            designerName: data.designerName || '',
            subSalesperson: data.subSalesperson || '',
            measurementEngineer: data.measurementEngineer || '',
            date: data.date || '',
            total: data.total || 0,
            discount: data.discount || 0,
            items: data.items || [],
            fees: data.fees || {},
            customFees: data.customFees || [],
            depositRecords: data.depositRecords || [],
            feeRemarks: data.feeRemarks || '',
            lastSynced: data.lastSynced || file.getLastUpdated().toISOString(),
            lastModified: file.getLastUpdated().toISOString(),
            currentUnit: data.currentUnit || 'mm'
          });
        } catch (e) {}
      }
    }
    
    quotes.sort((a, b) => new Date(b.lastModified) - new Date(a.lastModified));
    return { success: true, quotes };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function getQuote(username, projNo) {
  try {
    initializeFolders();
    const userFolder = getUserFolder(username);
    const files = userFolder.getFilesByName(projNo + '.json');
    
    if (files.hasNext()) {
      const file = files.next();
      return { success: true, quote: JSON.parse(file.getBlob().getDataAsString()) };
    }
    return { success: false, message: 'Quote not found' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function deleteQuote(username, projNo) {
  try {
    initializeFolders();
    const userFolder = getUserFolder(username);
    const files = userFolder.getFilesByName(projNo + '.json');
    while (files.hasNext()) files.next().setTrashed(true);
    logActivity(username, 'delete_quote', projNo, 'Deleted quote');
    return { success: true, message: 'Quote deleted' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// Get quote detail by projNo (searches all user folders)
function getQuoteDetail(projNo) {
  try {
    initializeFolders();
    const subFolders = mainFolder.getFolders();
    
    while (subFolders.hasNext()) {
      const folder = subFolders.next();
      if (folder.getName() === 'admin' || folder.getName() === 'squirrel analysis') continue;
      
      const files = folder.getFilesByName(projNo + '.json');
      if (files.hasNext()) {
        const file = files.next();
        const quote = JSON.parse(file.getBlob().getDataAsString());
        quote.createdBy = folder.getName();
        return { success: true, quote: quote };
      }
    }
    
    return { success: false, message: 'Quote not found' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// ========================================
// 佣金保存功能 (新增)
// ========================================

function saveQuoteCommission(projNo, salespersonCommission, subSalespersonCommission, measurementEngineerCommission, designerCommission) {
  try {
    initializeFolders();
    const subFolders = mainFolder.getFolders();
    
    while (subFolders.hasNext()) {
      const folder = subFolders.next();
      if (folder.getName() === 'admin' || folder.getName() === 'squirrel analysis') continue;
      
      const files = folder.getFilesByName(projNo + '.json');
      if (files.hasNext()) {
        const file = files.next();
        const quote = JSON.parse(file.getBlob().getDataAsString());
        
        // 更新佣金数据
        quote.salespersonCommission = parseFloat(salespersonCommission) || 0;
        quote.subSalespersonCommission = parseFloat(subSalespersonCommission) || 0;
        quote.measurementEngineerCommission = parseFloat(measurementEngineerCommission) || 0;
        quote.designerCommission = parseFloat(designerCommission) || 0;
        
        // 删除旧文件并保存新文件
        const fileName = projNo + '.json';
        const allFiles = folder.getFilesByName(fileName);
        while (allFiles.hasNext()) {
          allFiles.next().setTrashed(true);
        }
        folder.createFile(fileName, JSON.stringify(quote, null, 2), 'application/json');
        
        // 同时更新 analysis 文件夹中的已完成报价
        updateAnalysisCommission(projNo, quote);
        
        logActivity('admin', 'save_commission', projNo, 'Saved commission for quote: ' + projNo);
        
        return { success: true, message: 'Commission saved successfully' };
      }
    }
    
    return { success: false, message: 'Quote not found' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// 更新 analysis 文件夹中的已完成报价的佣金数据
function updateAnalysisCommission(projNo, quoteData) {
  try {
    const files = analysisFolder.getFiles();
    while (files.hasNext()) {
      const file = files.next();
      const fileName = file.getName();
      // 查找包含此 projNo 的已完成报价
      if (fileName.includes('_' + projNo + '.json') || fileName.includes('_' + projNo + '_')) {
        const analysisQuote = JSON.parse(file.getBlob().getDataAsString());
        // 更新佣金数据
        analysisQuote.salespersonCommission = quoteData.salespersonCommission;
        analysisQuote.subSalespersonCommission = quoteData.subSalespersonCommission;
        analysisQuote.measurementEngineerCommission = quoteData.measurementEngineerCommission;
        analysisQuote.designerCommission = quoteData.designerCommission;
        // 保存更新
        file.setTrashed(true);
        analysisFolder.createFile(fileName, JSON.stringify(analysisQuote, null, 2));
        break;
      }
    }
  } catch (e) {
    console.error('Failed to update analysis commission:', e);
  }
}

// ========================================
// 佣金保存功能结束
// ========================================

function completeQuote(username, projNo) {
  try {
    initializeFolders();
    const userFolder = getUserFolder(username);
    const files = userFolder.getFilesByName(projNo + '.json');
    
    if (!files.hasNext()) return { success: false, message: 'Quote not found' };
    
    const file = files.next();
    const quote = JSON.parse(file.getBlob().getDataAsString());
    quote.status = 'completed';
    quote.completedAt = new Date().toISOString();
    quote.completedBy = username;
    
    // Save to squirrel analysis folder with format: final_username_projNo.json
    const finalFileName = 'final_' + username + '_' + projNo + '.json';
    analysisFolder.createFile(finalFileName, JSON.stringify(quote, null, 2));
    logActivity(username, 'complete_quote', projNo, 'Completed quote: ' + (quote.customerName || projNo));
    return { success: true, quote, fileName: finalFileName };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function shareQuote(username, targetUsername, projNo) {
  try {
    initializeFolders();
    let quote = null;
    let sourceFolder = null;
    
    const subFolders = mainFolder.getFolders();
    while (subFolders.hasNext()) {
      const folder = subFolders.next();
      if (folder.getName() === 'admin' || folder.getName() === 'squirrel analysis' || folder.getName() === 'Commission') continue;
      
      const files = folder.getFilesByName(projNo + '.json');
      if (files.hasNext()) {
        const file = files.next();
        quote = JSON.parse(file.getBlob().getDataAsString());
        sourceFolder = folder;
        break;
      }
    }
    
    if (!quote || !sourceFolder) return { success: false, message: 'Quote not found' };
    
    const sharedFileName = 'shared_' + sourceFolder.getName() + '_' + projNo + '.json';
    quote.projNo = sharedFileName.replace('.json', '');
    quote.sharedFrom = username;
    quote.sharedAt = new Date().toISOString();
    
    const toFolder = getUserFolder(targetUsername);
    toFolder.createFile(sharedFileName, JSON.stringify(quote, null, 2));
    logActivity(username, 'share_quote', quote.projNo, 'Shared quote to ' + targetUsername);
    return { success: true, newProjNo: quote.projNo, message: 'Quote shared with ' + targetUsername };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// ==================== 产品数据管理 ====================

function getProducts() {
  try {
    initializeFolders();
    const files = adminFolder.getFilesByName(CONFIG.PRODUCTS_FILE);
    
    if (!files.hasNext()) {
      return { success: false, message: 'Products file not found. Please upload products.json to the admin folder.' };
    }
    
    const file = files.next();
    const products = JSON.parse(file.getBlob().getDataAsString());
    return { success: true, products: products };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function saveProducts(products) {
  try {
    initializeFolders();
    
    const files = adminFolder.getFilesByName(CONFIG.PRODUCTS_FILE);
    while (files.hasNext()) {
      files.next().setTrashed(true);
    }
    
    adminFolder.createFile(CONFIG.PRODUCTS_FILE, JSON.stringify(products, null, 2));
    productsCache = products;
    
    return { success: true, message: 'Products saved successfully' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function getAnalysisData() {
  try {
    initializeFolders();
    const quotes = [];
    const files = analysisFolder.getFiles();
    
    while (files.hasNext()) {
      const file = files.next();
      if (file.getName().startsWith('final_')) {
        try { quotes.push(JSON.parse(file.getBlob().getDataAsString())); } catch (e) {}
      }
    }
    
    quotes.sort((a, b) => new Date(b.completedAt) - new Date(a.completedAt));
    return { success: true, quotes };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function getAllQuotes() {
  try {
    initializeFolders();
    const quotes = [];
    const subFolders = mainFolder.getFolders();
    
    while (subFolders.hasNext()) {
      const folder = subFolders.next();
      if (folder.getName() === 'admin' || folder.getName() === 'squirrel analysis') continue;
      
      const files = folder.getFiles();
      while (files.hasNext()) {
        const file = files.next();
        if (file.getName().endsWith('.json') && !file.getName().startsWith('final_')) {
          try {
            const quote = JSON.parse(file.getBlob().getDataAsString());
            quote.createdBy = folder.getName();
            quotes.push(quote);
          } catch (e) {}
        }
      }
    }
    
    quotes.sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));
    return { success: true, quotes };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// ==================== 佣金管理 (分开储存版本) ====================

// 获取单个角色的佣金数据
function getCommission(projNo, role) {
  try {
    initializeFolders();
    if (!role) {
      // 如果没有指定role，返回所有4个角色的佣金
      return getAllCommissions(projNo);
    }
    
    const fileName = 'comm_' + role + '_' + projNo + '.json';
    const files = commissionFolder.getFilesByName(fileName);
    
    if (files.hasNext()) {
      const file = files.next();
      const commission = JSON.parse(file.getBlob().getDataAsString());
      return { success: true, commission: commission };
    }
    
    return { success: true, commission: null };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// 获取所有4个角色的佣金数据
function getAllCommissions(projNo) {
  try {
    initializeFolders();
    const roles = ['销售员', '副销售员', '测量工程师', '设计师'];
    const result = {
      projNo: projNo,
      grandTotal: 0,
      salesperson: { name: '', percentage: 0, amount: 0 },
      subSalesperson: { name: '', percentage: 0, amount: 0 },
      measurementEngineer: { name: '', percentage: 0, amount: 0 },
      designer: { name: '', percentage: 0, amount: 0 }
    };
    
    for (const role of roles) {
      const fileName = 'comm_' + role + '_' + projNo + '.json';
      const files = commissionFolder.getFilesByName(fileName);
      
      if (files.hasNext()) {
        const file = files.next();
        const commission = JSON.parse(file.getBlob().getDataAsString());
        const roleKey = getRoleKey(role);
        if (roleKey && commission) {
          result[roleKey] = {
            name: commission.name || '',
            percentage: commission.percentage || 0,
            amount: commission.amount || 0
          };
          if (commission.grandTotal) result.grandTotal = commission.grandTotal;
        }
      }
    }
    
    return { success: true, commission: result };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// 保存单个角色的佣金数据
function saveCommission(projNo, role, name, percentage, grandTotal) {
  try {
    initializeFolders();
    const fileName = 'comm_' + role + '_' + projNo + '.json';
    
    const commissionData = {
      projNo: projNo,
      role: role,
      name: name || '',
      percentage: parseFloat(percentage) || 0,
      amount: (parseFloat(grandTotal) || 0) * (parseFloat(percentage) || 0) / 100,
      grandTotal: parseFloat(grandTotal) || 0,
      savedAt: new Date().toISOString()
    };
    
    // Delete existing file
    const files = commissionFolder.getFilesByName(fileName);
    while (files.hasNext()) {
      files.next().setTrashed(true);
    }
    
    // Create new file
    commissionFolder.createFile(fileName, JSON.stringify(commissionData, null, 2), 'application/json');
    
    return { success: true, message: 'Commission saved successfully' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// 辅助函数：根据角色名称获取key
function getRoleKey(role) {
  const roleMap = {
    '销售员': 'salesperson',
    '副销售员': 'subSalesperson',
    '测量工程师': 'measurementEngineer',
    '设计师': 'designer'
  };
  return roleMap[role] || null;
}

// =====================================================
// Quote Analysis 报价分析系统 - 新增 API
// =====================================================

/**
 * 列出 Squirrel Designer/ 下所有用户文件夹
 * action: list_user_folders
 */
function listUserFolders() {
  try {
    // 2026-08-15: 30s 跨调用 cache（CacheService 持久化）
    const cached = cacheGet('qa_folders_v1');
    if (cached) return { success: true, folders: cached, cached: true };

    initializeFolders();
    const folders = [];
    const subs = mainFolder.getFolders();
    while (subs.hasNext()) {
      const folder = subs.next();
      const name = folder.getName();
      if (name !== 'admin' && name !== 'squirrel analysis' && name !== 'Commission') {
        folders.push({ name, id: folder.getId() });
      }
    }
    folders.sort((a, b) => a.name.localeCompare(b.name));
    cachePut('qa_folders_v1', folders, CACHE_TTL_FOLDERS);
    folderListCache = { ts: Date.now(), folders };
    return { success: true, folders };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/**
 * 列出指定用户文件夹中的报价文件
 * action: list_user_quote_files
 * @param username - 用户文件夹名
 */
function listUserQuoteFiles(username) {
  try {
    // 2026-08-15: 缓存 user folder id（DriveApp Folder 对象本身不能跨调用 serialize）
    // 用 CacheService 存 id 字符串，getFolderById 是 O(1) lookup
    initializeFolders();
    let userFolder;
    const cachedId = cacheGet('qa_userFolder_' + username);
    if (cachedId) {
      try { userFolder = DriveApp.getFolderById(cachedId); } catch (e) { userFolder = null; }
    }
    if (!userFolder) {
      userFolder = getUserFolder(username);
      cachePut('qa_userFolder_' + username, userFolder.getId(), CACHE_TTL_USER_FOLDER);
    }
    userFolderCache[username] = { ts: Date.now(), folder: userFolder };
    const files = userFolder.getFiles();
    const fileList = [];

    while (files.hasNext()) {
      const file = files.next();
      const name = file.getName();
      if (name.endsWith('.json') && !name.startsWith('final_') && !name.startsWith('shared_')) {
        // 2026-08-15 优化：list 端点不需要完整 items[]，但 customerName/total 必须有
        // 用 getBlob() + parse 一次，只读全文（DriveApp 不支持部分读）
        // 加速点：去掉空字段、避免无谓的 reduce
        try {
          const data = JSON.parse(file.getBlob().getDataAsString());
          fileList.push({
            id: file.getId(),
            name: name,
            projNo: data.projNo || name.replace('.json', ''),
            customerName: data.customerName || '',
            totalSales: data.total || data.grandTotal || data.totalSales || 0,
            modifiedTime: file.getLastUpdated().toISOString()
          });
        } catch (e) {
          fileList.push({ id: file.getId(), name: name, projNo: name.replace('.json', ''), customerName: '', totalSales: 0, modifiedTime: file.getLastUpdated().toISOString() });
        }
      }
    }
    fileList.sort((a, b) => new Date(b.modifiedTime) - new Date(a.modifiedTime));
    return { success: true, files: fileList };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/**
 * 读取指定用户的报价文件
 * action: read_user_quote_file
 * @param username
 * @param fileId
 */
function readUserQuoteFile(username, fileId) {
  try {
    const file = DriveApp.getFileById(fileId);
    const content = file.getBlob().getDataAsString();
    const data = JSON.parse(content);
    return { success: true, fileId, fileName: file.getName(), data };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/**
 * 读取指定用户的报价文件（通过文件名）
 * action: read_user_quote_file_by_name
 * @param username
 * @param fileName
 */
function readUserQuoteFileByName(username, fileName) {
  try {
    initializeFolders();
    const userFolder = getUserFolder(username);
    const files = userFolder.getFilesByName(fileName);
    if (!files.hasNext()) return { success: false, message: 'File not found' };
    const file = files.next();
    return { success: true, fileId: file.getId(), fileName, data: JSON.parse(file.getBlob().getDataAsString()) };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/**
 * 导入报价：把原 JSON 完整拷贝到 squirrel analysis/ 作为 check_{projNo}.json
 * 不修改原文件，不污染 Squirrel Designer 数据
 * action: import_quote_to_analysis
 * @param username - 用户文件夹名（用于查找原文件）
 * @param fileId - 原 JSON 的 Drive 文件 ID
 */
function importQuoteToAnalysis(username, fileId) {
  try {
    initializeFolders();
    const file = DriveApp.getFileById(fileId);
    const fileName = file.getName();
    if (!fileName.endsWith('.json')) {
      return { success: false, message: 'Not a JSON file' };
    }
    const projNo = fileName.replace('.json', '');
    const raw = JSON.parse(file.getBlob().getDataAsString());

    // Add metadata fields (these are the only additions — original fields untouched)
    raw.projNo = projNo;
    raw.lastModified = new Date().toISOString().split('T')[0];
    raw.status = raw.status || 'pending';

    // Save as check_{projNo}.json in squirrel analysis/
    const checkFileName = 'check_' + projNo + '.json';
    const existing = analysisFolder.getFilesByName(checkFileName);
    while (existing.hasNext()) {
      existing.next().setTrashed(true);
    }
    // 2026-08-15: 去掉 null,2 — 文件小一半，parse 快一倍
    analysisFolder.createFile(checkFileName, JSON.stringify(raw), 'application/json');
    // 失效 list 缓存（CacheService 跨调用 + 模块级 fallback）
    cacheRemove('qa_list_v1');
    checkQuoteListCache = null;

    logActivity('system', 'import_quote_to_analysis', projNo, 'Imported from ' + username + '/' + fileName);
    return { success: true, message: 'Imported as ' + checkFileName, projNo, data: raw };
  } catch (error) {
    return { success: false, message: 'importQuoteToAnalysis: ' + error.toString() };
  }
}

/**
 * 保存报价分析数据到 squirrel analysis/ 文件夹
 * action: save_check_quote
 * @param projNo - 项目编号 (如 test00)
 * @param quoteData - 完整的报价分析数据
 */
function saveCheckQuote(projNo, quoteData) {
  try {
    initializeFolders();
    const fileName = 'check_' + projNo + '.json';

    // Defensive: if quoteData is a string (double-encoded from frontend), unwrap once
    if (typeof quoteData === 'string') {
      try { quoteData = JSON.parse(quoteData); } catch (e) { /* keep as-is */ }
    }
    if (!quoteData || typeof quoteData !== 'object') {
      return { success: false, message: 'Invalid quoteData payload' };
    }
    
    // 添加保存时间戳
    quoteData.lastSaved = new Date().toISOString();

    // 删除旧文件
    const files = analysisFolder.getFilesByName(fileName);
    while (files.hasNext()) {
      files.next().setTrashed(true);
    }

    // 2026-08-15: 去掉 null,2 — 文件小一半，parse 快一倍
    analysisFolder.createFile(fileName, JSON.stringify(quoteData), 'application/json');
    // 失效 list 缓存（CacheService 跨调用持久 + 模块级 fallback）
    cacheRemove('qa_list_v1');
    checkQuoteListCache = null;

    logActivity('system', 'save_check_quote', projNo, 'Saved check quote: ' + projNo);

    return { success: true, message: 'Quote saved to analysis folder', fileName: fileName };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/**
 * 清空 squirrel analysis/ 文件夹里所有 check_*.json
 * action: clean_analysis_folder
 * DEBUG ONLY: 用来清掉双重 JSON 编码的脏数据
 */
function cleanAnalysisFolder() {
  try {
    initializeFolders();
    const files = analysisFolder.getFiles();
    const removed = [];
    while (files.hasNext()) {
      const file = files.next();
      const name = file.getName();
      if (name.startsWith('check_') && name.endsWith('.json')) {
        file.setTrashed(true);
        removed.push(name);
      }
    }
    cacheRemove('qa_list_v1');
    checkQuoteListCache = null;  // 失效 list 缓存
    logActivity('system', 'clean_analysis_folder', '-', 'Removed ' + removed.length + ' files: ' + removed.join(', '));
    return { success: true, message: 'Removed ' + removed.length + ' files', removed: removed };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/**
 * 轻量级单报价元数据查询（不返回 items[] 等大数据）
 * 用于 list 视图快速刷新某个 quote 的状态
 * action: get_check_quote_meta
 * @param projNo
 */
function getCheckQuoteMeta(projNo) {
  try {
    initializeFolders();
    const fileName = 'check_' + projNo + '.json';
    const files = analysisFolder.getFilesByName(fileName);
    if (!files.hasNext()) {
      return { success: false, message: 'Quote not found: ' + projNo };
    }
    const file = files.next();
    let data = JSON.parse(file.getBlob().getDataAsString());
    if (typeof data === 'string') {
      try { data = JSON.parse(data); } catch (e) { /* keep as-is */ }
    }
    // 只返回 list 视图需要的字段
    return {
      success: true,
      meta: {
        projNo: data.projNo || projNo,
        customerName: data.customerName || data.customer?.name || '',
        totalSales: data.totalSales || 0,
        totalCost: data.totalCost || 0,
        profit: data.profit || 0,
        profitPercent: data.profitPercent || 0,
        debtPercent: data.debtPercent || 0,
        status: data.status || 'pending',
        lastModified: data.lastModified || data.lastSaved || file.getLastUpdated().toISOString(),
        fileName: fileName
      }
    };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/**
 * 列出 squirrel analysis/ 文件夹中所有已保存的报价分析
 * action: get_check_quote_list
 */
function getCheckQuoteList() {
  try {
    // 2026-08-15: 10s 跨调用 cache（CacheService 持久化）
    const cached = cacheGet('qa_list_v1');
    if (cached) return { success: true, quotes: cached, cached: true };

    initializeFolders();
    const files = analysisFolder.getFiles();
    const quotes = [];

    while (files.hasNext()) {
      const file = files.next();
      const name = file.getName();
      // 只返回 check_ 开头的 JSON 文件
      if (name.startsWith('check_') && name.endsWith('.json')) {
        try {
          let data = JSON.parse(file.getBlob().getDataAsString());
          // Defensive: handle double-encoded legacy data
          if (typeof data === 'string') {
            try { data = JSON.parse(data); } catch (e) { /* keep as-is */ }
          }
          quotes.push({
            projNo: data.projNo || name.replace('check_', '').replace('.json', ''),
            customerName: data.customerName || data.customer?.name || '',
            totalSales: data.totalSales || 0,
            totalCost: data.totalCost || 0,
            profit: data.profit || 0,
            profitPercent: data.profitPercent || 0,
            debtPercent: data.debtPercent || 0,
            status: data.status || 'pending',
            lastModified: data.lastModified || data.lastSaved || file.getLastUpdated().toISOString(),
            fileName: name
          });
        } catch (e) {
          // 跳过无效的 JSON 文件
        }
      }
    }

    // 按修改时间倒序排列
    quotes.sort((a, b) => new Date(b.lastModified) - new Date(a.lastModified));
    cachePut('qa_list_v1', quotes, CACHE_TTL_LIST);
    checkQuoteListCache = { ts: Date.now(), quotes };
    return { success: true, quotes: quotes };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/**
 * 读取单个报价分析的完整数据
 * action: get_check_quote
 * @param projNo - 项目编号
 */
function getCheckQuote(projNo) {
  try {
    initializeFolders();
    const fileName = 'check_' + projNo + '.json';
    const files = analysisFolder.getFilesByName(fileName);

    if (!files.hasNext()) {
      return { success: false, message: 'Quote not found: ' + projNo };
    }

    const file = files.next();
    let data = JSON.parse(file.getBlob().getDataAsString());

    // Defensive: if data is a string (double-encoded from legacy saves), unwrap once
    if (typeof data === 'string') {
      try { data = JSON.parse(data); } catch (e) { /* keep as-is */ }
    }

    return { success: true, quote: data };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/**
 * 删除已保存的报价分析
 * action: delete_check_quote
 * @param projNo - 项目编号
 */
function deleteCheckQuote(projNo) {
  try {
    initializeFolders();
    const fileName = 'check_' + projNo + '.json';
    const files = analysisFolder.getFilesByName(fileName);
    
    if (!files.hasNext()) {
      return { success: false, message: 'Quote not found: ' + projNo };
    }
    
    while (files.hasNext()) {
      files.next().setTrashed(true);
    }

    cacheRemove('qa_list_v1');
    checkQuoteListCache = null;  // 失效 list 缓存
    logActivity('system', 'delete_check_quote', projNo, 'Deleted check quote: ' + projNo);

    return { success: true, message: 'Quote deleted' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// =====================================================
// API 入口
// =====================================================

function doGet(e) {
  try {
    const action = e.parameter.action;
    let result;
    
    switch (action) {
      case 'login':
        result = login(e.parameter.username, e.parameter.password);
        break;
      case 'saveQuote':
        const quoteData = JSON.parse(e.parameter.quoteData || '{}');
        result = saveQuote(e.parameter.username, e.parameter.projNo, quoteData);
        break;
      case 'getQuoteList':
        result = getQuoteList(e.parameter.username);
        break;
      case 'getQuote':
        result = getQuote(e.parameter.username, e.parameter.projNo);
        break;
      case 'deleteQuote':
        result = deleteQuote(e.parameter.username, e.parameter.projNo);
        break;
      case 'completeQuote':
        result = completeQuote(e.parameter.username, e.parameter.projNo);
        break;
      case 'shareQuote':
        result = shareQuote(e.parameter.username, e.parameter.targetUsername, e.parameter.projNo);
        break;
      case 'adminLogin':
        result = adminLogin(e.parameter.username, e.parameter.password);
        break;
      case 'changeAdminPassword':
        result = changeAdminPassword(e.parameter.oldPassword, e.parameter.newPassword);
        break;
      case 'getAdminStats':
        result = getAdminStats();
        break;
      case 'getDetailedUsers':
        result = getDetailedUsers();
        break;
      case 'getUserByUsername':
        result = getUserByUsername(e.parameter.username);
        break;
      case 'updateUser':
        let updateData;
        if (e.parameter.updates) {
          updateData = JSON.parse(e.parameter.updates);
        } else {
          updateData = {};
          if (e.parameter.displayName) updateData.displayName = e.parameter.displayName;
          if (e.parameter.password) updateData.password = e.parameter.password;
          if (e.parameter.departments) updateData.departments = e.parameter.departments.split(',').filter(d => d);
          if (e.parameter.isActive !== undefined) updateData.isActive = e.parameter.isActive === 'true';
        }
        result = updateUser(e.parameter.username, updateData);
        break;
      case 'adminAddUser':
        result = adminAddUser(e.parameter.username, e.parameter.password, e.parameter.displayName, JSON.parse(e.parameter.departments || '[]'));
        break;
      case 'adminDeleteUser':
        result = adminDeleteUser(e.parameter.username);
        break;
      case 'adminToggleUserStatus':
        result = adminToggleUserStatus(e.parameter.username);
        break;
      case 'getUserActivities':
        result = getUserActivities(e.parameter.username);
        break;
      case 'getUserStatistics':
        result = getUserStatistics(e.parameter.username);
        break;
      case 'getQuoteDetail':
        result = getQuoteDetail(e.parameter.projNo);
        break;
      case 'getAllActivities':
        result = getAllActivities();
        break;
      case 'getAnalysisData':
        result = getAnalysisData();
        break;
      case 'getAllQuotes':
        result = getAllQuotes();
        break;
      case 'initialize':
        result = initializeFolders();
        break;
      case 'getProducts':
        result = getProducts();
        break;
      case 'getUsers':
        initializeFolders();
        result = { success: true, users: usersCache.users.map(u => ({ username: u.username, displayName: u.displayName || u.username })) };
        break;
      case 'saveProducts':
        result = saveProducts(JSON.parse(e.parameter.products || '[]'));
        break;
      // 佣金API (新增支持role参数)
      case 'getCommission':
        result = getCommission(e.parameter.projNo, e.parameter.role);
        break;
      case 'saveCommission':
        result = saveCommission(
          e.parameter.projNo,
          e.parameter.role,
          e.parameter.name || '',
          e.parameter.percentage || 0,
          e.parameter.grandTotal || 0
        );
        break;
      case 'getAllCommissions':
        result = getAllCommissions(e.parameter.projNo);
        break;
      // ===== Quote Analysis API (新增) =====
      case 'list_user_folders':
        result = listUserFolders();
        break;
      case 'list_user_quote_files':
        result = listUserQuoteFiles(e.parameter.username);
        break;
      case 'read_user_quote_file':
        result = readUserQuoteFile(e.parameter.username, e.parameter.fileId);
        break;
      case 'read_user_quote_file_by_name':
        result = readUserQuoteFileByName(e.parameter.username, e.parameter.fileName);
        break;
      case 'import_quote_to_analysis':
        result = importQuoteToAnalysis(e.parameter.username, e.parameter.fileId);
        break;
      case 'save_check_quote':
        result = saveCheckQuote(e.parameter.projNo, JSON.parse(e.parameter.quoteData || '{}'));
        break;
      case 'get_check_quote_list':
        result = getCheckQuoteList();
        break;
      case 'clean_analysis_folder':
        result = cleanAnalysisFolder();
        break;
      case 'get_check_quote':
        result = getCheckQuote(e.parameter.projNo);
        break;
      case 'get_check_quote_meta':
        result = getCheckQuoteMeta(e.parameter.projNo);
        break;
      case 'delete_check_quote':
        result = deleteCheckQuote(e.parameter.projNo);
        break;
      case 'ping':
        result = { success: true, message: 'API is running', timestamp: new Date().toISOString() };
        break;
      default:
        result = { success: false, message: 'Unknown action: ' + action };
    }
    
    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: error.toString() })).setMimeType(ContentService.MimeType.JSON);
  }
}

// POST 请求入口：支持 save_check_quote 传大 quoteData（通过 body）
function doPost(e) {
  try {
    let payload;
    const contentType = e.postData ? e.postData.type : '';
    
    if (contentType && contentType.includes('application/json')) {
      // ContentService JSON POST
      payload = JSON.parse(e.postData.contents);
    } else if (e.postData && e.postData.contents) {
      // text/plain 或其他格式，直接取 contents
      try {
        payload = JSON.parse(e.postData.contents);
      } catch (parseErr) {
        // 如果不是 JSON，当作 key=value&key2=value2 格式解析
        payload = {};
        const parts = e.postData.contents.split('&');
        for (const part of parts) {
          const [k, v] = part.split('=');
          if (k) payload[decodeURIComponent(k)] = decodeURIComponent(v || '');
        }
      }
    } else {
      return ContentService.createTextOutput(JSON.stringify({ success: false, message: 'No post data received' })).setMimeType(ContentService.MimeType.JSON);
    }
    
    const action = payload.action;
    let result;
    
    switch (action) {
      case 'save_check_quote':
        // POST 方式：projNo 和 quoteData 从 body 取
        result = saveCheckQuote(payload.projNo, payload.quoteData || {});
        break;
      case 'import_quote_to_analysis':
        result = importQuoteToAnalysis(payload.username, payload.fileId);
        break;
      case 'ping':
        result = { success: true, message: 'API is running (POST)', timestamp: new Date().toISOString() };
        break;
      default:
        // 其他 action 也支持 POST（把 payload 展开到 parameter 格式）
        result = doGet({ parameter: payload });
    }
    
    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: error.toString() })).setMimeType(ContentService.MimeType.JSON);
  }
}